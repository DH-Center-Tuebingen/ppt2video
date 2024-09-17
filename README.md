# ppt2video

Converts PowerPoint file to video with text-to-speech generated audio from slide notes.

**Conversion workflow**:

- Extract presentation slides as images.
- Generate audio from slide notes using either the Microsoft Speech API (medium quality voices, pre-installed with Windows) or Microsoft Azure Speech Services (high-quality voices, requires a subscription and API key).
- Merge audio and images to create a video presentation.

**Requirements**:

- Windows
- Python 3 with `pywin32` installed
- PowerPoint installed
- ffmpeg installed
- If Azure Speech Services shall be used for speech synthesis, the Python package `azure-cognitiveservices-speech` and an Azure subscription and Speech Services API key are required

Licensed under the [MIT License](./LICENSE).