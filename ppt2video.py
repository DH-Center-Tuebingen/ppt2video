import argparse
import os
import re
import subprocess
from tempfile import mkdtemp
import win32com.client as win32

# Make full path from relative path for command line arguments
def ensure_full_path(s):
    return os.path.abspath(s) if not os.path.isabs(s) else s

# Parse command line arguments
parser = argparse.ArgumentParser(description='Convert PowerPoint presentation to video presentation with synthesized audio.')
parser.add_argument('pptfile', help='PowerPoint file to create presentation video for', type=ensure_full_path)
parser.add_argument('output', help='Output video file')
parser.add_argument('--slides', help='Slide numbers or ranges to extract (comma-separated list)', type=str)
parser.add_argument('--silence', default=1.5, help='Seconds of silence to pad each slide audio; default: 1.5', type=float)
parser.add_argument('--voice', default='en-GB-SoniaNeural', help='Voice identifier. When using Azure see https://learn.microsoft.com/en-us/azure/ai-services/speech-service/language-support?tabs=tts#prebuilt-neural-voices (default: en-GB-SoniaNeural); when using SAPI, use index number of voice installed on your system')
parser.add_argument('--pronunciation_mapping', help="File that maps the spelling of words and acronyms to their pronunciation", type=ensure_full_path)
parser.add_argument('--video_width', default=1920, help='Width of the output video in pixels; default: 1920', type=int)
parser.add_argument('--video_height', default=1080, help='Height of the output video in pixels; default: 1080', type=int)
parser.add_argument('--api', default='Azure', choices=['SAPI', 'Azure'], help='API to use for speech synthesis: Azure (default; Microsoft Azure AI Speech SDK, requires API key set as environment variable SPEECH_KEY and region in SPEECH_REGION) or SAPI (Microsoft Speech API, part of Windows)', type=str)
parser.add_argument('--update', type=ensure_full_path, help='Folder with temporary files from previous conversion to reuse when only a subset of the slides were updated; should be used together with --slides, must use the same file extension for the output file (i.e., the same video container format) as the previous conversion, and the slide order and count must be the same as in the previous conversion')
parser.add_argument('--quit_ppt', action='store_true', help='Quit PowerPoint after processing the presentation, if no other presentations are open')
args = parser.parse_args()

# import API
if args.api == 'Azure':
    import azure.cognitiveservices.speech as speechsdk

# Pronunciation mapping file contains lines matching the pattern word=pronunciation, e.g.:
# FDAT=effdutt
# Load the file and store in dictionary
pronunciation_mapping = {}
if args.pronunciation_mapping:
    with open(args.pronunciation_mapping, 'r', encoding="utf-8") as f:
        for line in f:
            word, pronunciation = line.strip().split('=')
            pronunciation_mapping[word.lower().strip()] = pronunciation.lower().strip()

# Init Azure Speech SDK
if args.api == 'Azure':
    speech_config = speechsdk.SpeechConfig(
        subscription=os.environ.get('SPEECH_KEY'), 
        region=os.environ.get('SPEECH_REGION'))
    speech_config.speech_synthesis_voice_name = args.voice

# Determine video container format from output file extension
container_format = os.path.splitext(args.output)[1][1:]

# Create temp directories
temp_dir = args.update if args.update else mkdtemp()
slide_folder = os.path.join(temp_dir, "slides")
audio_folder = os.path.join(temp_dir, "audio")
video_folder = os.path.join(temp_dir, "video")
if not args.update:
    for dir in [slide_folder, audio_folder, video_folder]:
        os.mkdir(dir)

# Open PowerPoint file, will be needed to extract slide images and notes
ppt = win32.Dispatch("PowerPoint.Application")
presentation = ppt.Presentations.Open(args.pptfile, True, False, False)

# Extract desired slide list
slide_list = []
if args.slides is None:
    slide_list = range(1, presentation.Slides.Count + 1)        
else:
    slides = [s.strip() for s in args.slides.split(',')]
    for i in slides:
        if '-' in i:
            start, end = i.split('-')
            slide_list.extend(range(int(start), int(end) + 1))
        else:
            slide_list.append(int(i))
    slide_list = sorted(slide_list)

# Generate silence file
print(f"Generating {args.silence} seconds silence audio file")
silence_file = os.path.join(temp_dir, "silence.wav")
subprocess.run(f'ffmpeg -y -hide_banner -loglevel error -f lavfi -i anullsrc=r=11025:cl=mono -t {args.silence} -c:a pcm_s16le {silence_file}', shell=True)
    
# remember created slide videos for ffmpeg concat later
slide_videos = []
total_chars = 0

# Loop over slides in presentation
for slide_number in slide_list:
    print(f"Processing slide {slide_number}")
    slide = presentation.Slides(slide_number)

    # Read slide text
    slide_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text    
    if not slide_text or len(slide_text.strip()) == 0:
        print(f"  Skipping slide {slide_number} because no note text found")
        continue
    
    # Remove newlines and carriage returns
    slide_text = slide_text.replace('\n', ' ').replace('\r', ' ').strip()
    
    # Replace words with pronunciation from mapping file
    for word, pronunciation in pronunciation_mapping.items():
        slide_text = re.sub(rf'\b{re.escape(word)}\b', pronunciation, slide_text, flags=re.IGNORECASE)
    
    total_chars += len(slide_text)
    
    # Export slide as image
    print(f"  Exporting slide {slide_number} as image")
    slide_image_file = os.path.join(slide_folder, f"slide_{slide_number}.png")
    if os.path.exists(slide_image_file):
        os.remove(slide_image_file)
    slide.Export(slide_image_file, "PNG", ScaleWidth=args.video_width, ScaleHeight=args.video_height)
        
    # Synthesize audio for slide text
    audio_file = os.path.join(audio_folder, f"audio_{slide_number}.wav")
    print(f"  Synthesizing audio of slide {slide_number}")    
    if args.api == 'Azure': 
        audio_config = speechsdk.audio.AudioOutputConfig(filename=audio_file)
        speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
        speech_synthesis_result = speech_synthesizer.speak_text_async(slide_text).get()
        if speech_synthesis_result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            pass
        elif speech_synthesis_result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = speech_synthesis_result.cancellation_details
            print(f"  Speech synthesis canceled: {cancellation_details.reason}")
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                if cancellation_details.error_details:
                    print(f"  Error details: {cancellation_details.error_details}. Did you set the speech resource key and region values?")
            break
    else: # SAPI
        sapi = win32.Dispatch("SAPI.SpVoice")
        sapi.Voice = sapi.GetVoices().Item(int(args.voice))
        outfile = win32.Dispatch("SAPI.SpFileStream")
        outfile.Open(audio_file, 3, False)
        sapi.AudioOutputStream = outfile
        sapi.Speak(slide_text)
        outfile.Close()
    
    # Extend audio with silence in front and back and encode using AAC codec
    audio_file_padded = audio_file.replace('.wav', '.m4a')
    subprocess.run('ffmpeg -y -hide_banner -loglevel error -i "{silence}" -i "{audio_in}" -i "{silence}" -filter_complex "[0:0][1:0][2:0]concat=n=3:v=0:a=1[a]" -map "[a]" -c:a aac -strict experimental "{audio_out}"'.format(silence=silence_file, audio_in=audio_file, audio_out=audio_file_padded), shell=True)
    
    # Create video from slide image and synthesized audio
    video_file = os.path.join(video_folder, f"video_{slide_number}.{container_format}")
    slide_videos.append(video_file)
    
    print(f"  Creating video of slide {slide_number} with synthesized audio and slide image")
    subprocess.run('ffmpeg -y -hide_banner -loglevel error -loop 1 -i "{slide}" -i "{audio}" -c:v libx264 -framerate 5 -c:a copy -tune stillimage -shortest {video}'.format(slide=os.path.join(slide_folder, slide_image_file), audio=audio_file_padded, video=video_file), shell=True)

# Close PowerPoint
presentation.Close()

# If no open presentations remaining - quit PowerPoint
if args.quitppt and ppt.Presentations.Count == 0:
    ppt.Quit()

if len(slide_videos) == 0:
    print(f"No slide videos {'updated' if args.update else 'created'}, exiting")
    exit(0)

# Create final video by concatenating all slide videos
concat_file = os.path.join(temp_dir, "concat.txt")
# Create list of slide videos to concatenate only if not updating previous conversion
if not args.update:
    with open(concat_file, "w") as video_list_file:
        for slide_video in slide_videos:
            video_list_file.write(f"file '{slide_video}'\n")

print("Creating full video by concatenating all slide videos")
subprocess.run(f'ffmpeg -y -hide_banner -loglevel error -f concat -safe 0 -i "{concat_file}" -c copy "{args.output}"', shell=True)

print(f"Total characters synthesized: {total_chars}")
print(f"Temporary files kept in {temp_dir}")
print("Done.")