## Video Transcription & Presentation Generation Pipeline

This project provides an end-to-end solution to process video files for transcription and generate a presentation based on the analysis. It downloads a video, processes the audio (extraction, speed-up, segmentation), transcribes audio segments using OpenAI's Whisper model, asks for segment wise (not the audio segments, rather LLM constructed logical units) summary from an LLM followed by a Q and A and finally creates a PowerPoint presentation (with a PDF conversion option) that arranges each segment (logical) neatly together. Example output is uploaded for the url: https://www.youtube.com/live/D7BzTxVVMuw?t=19684s

#### Features

- **Video Download & Processing:**
  - Download videos (using yt-dlp) in lower resolution (e.g., 480p) to save bandwidth.
  - Extract screenshots at defined intervals.
  - Extract audio from the video and speed it up (e.g., 1.2x) for faster transcription.
  - Split audio into 1-minute segments for granular processing.

- **Transcription & Analysis:**
  - Transcribe each audio segment using the Whisper model.
  - Generate structured JSON outputs per segment containing:
    - **Summary:** A concise overview of the segment.
    - **Justification:** Why the segment is important.
    - **Counter-Intuitive Points:** Surprising insights worth noting.
    - **Topics:** Tags or topics for further exploration.
    - **Q&A Pairs:** Questions with concise answers to test understanding.
  - Each JSON output includes a `segment_number` and a `short_title`.

- **Presentation Generation:**
  - Create a PowerPoint presentation with 3 slides per segment:
    1. **Slide 1:** Title slide showing the segment’s `short_title`.
    2. **Slide 2:** Summary slide covering the segment summary, counter-intuitive points, and topics.
    3. **Slide 3:** Q&A slide containing the question–answer pairs.
  - Automatically convert the presentation to a PDF (using LibreOffice in headless mode).

- **Version Control:**
  - Example instructions are provided for pushing the local repository to GitHub.

#### Requirements

- **Programming Language:** Python 3.x  
- **Key Python Packages:**
  - [yt-dlp](https://github.com/yt-dlp/yt-dlp)
  - [FFmpeg](https://ffmpeg.org/) (must be installed separately)
  - [OpenAI Whisper](https://github.com/openai/whisper)
  - [python-pptx](https://python-pptx.readthedocs.io/)
- **Additional Tools:**
  - LibreOffice (for PPTX to PDF conversion; ensure it's installed and in your system's PATH)
  - Standard libraries: `subprocess`, `os`, `json`, `glob`, `textwrap`, etc.

#### Installation

1. **Clone the Repository:**

```bash
   git clone https://github.com/yourusername/yourrepo.git
   cd yourrepo
```

2. **Install Python Dependencies:**

```bash
pip install yt-dlp
pip install git+https://github.com/openai/whisper.git
pip install python-pptx
```

3. **Install FFmpeg & LibreOffice:**

For Ubuntu:

```bash
sudo apt-get update
sudo apt-get install ffmpeg libreoffice
```

## Usage

#### Video Processing & Transcription
Configure Your Video URL:
Update the video_url variable in the main script to point to your desired video.

Run the Pipeline:
The .ipynb file has code to:

- Download the video.
- Extract and speed up audio.
- Split audio into segments.
- Transcribe each segment using Whisper.
- Generate JSON outputs for segment summaries and Q&A pairs.

#### Presentation Generation

Generate the PPTX:
The code reads the JSON outputs (segment summaries and Q&A pairs) and creates a PowerPoint presentation:
Slide 1: Displays the segment's short_title.
Slide 2: Contains the summary, counter-intuitive points, and topics.
Slide 3: Lists the Q&A pairs.
Convert to PDF:
The script attempts to convert the PPTX to a PDF using LibreOffice in headless mode. Ensure LibreOffice is installed and available in the system's PATH.

## Contributing

Contributions are welcome! 
Please open an issue or submit a pull request to discuss improvements or new features.

## Licence

This project is licensed under the MIT License. See the LICENSE file for details.


