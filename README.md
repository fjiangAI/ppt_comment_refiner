# 🎨 PPT Comment Refiner & Audio Automation

Enhance your PowerPoint presentation notes effortlessly and add voiceovers automatically using OpenAI's GPT and TTS models.

## 🔥 Recent Updates

- **2024-11-17**: Added automated audio narration feature using OpenAI's TTS model, allowing users to generate voiceovers based on presentation notes. (only support for windows.)
- **2024-11-09**: Added automated comment refining using OpenAI's TTS model.

## 📋 Overview

PPT Comment Refiner & Audio Automation is a Python tool designed to refine and enhance the notes in your PowerPoint presentations and optionally add automated voiceovers to each slide. By leveraging OpenAI's GPT and TTS models, this tool improves the quality of your presentation notes and provides audio narration, making them more suitable for academic, professional, and even multimedia settings.

## ✨ Features

- **Automated Note Enhancement**: Utilizes OpenAI's GPT model to refine presentation notes for better clarity and eloquence.
- **Audio Narration Generation**: Automatically generates audio voiceovers based on the presentation notes using OpenAI's TTS model.
- **Seamless Integration**: Processes existing PowerPoint files and outputs refined and enhanced versions without manual intervention.
- **User-Friendly**: Simple setup and execution with minimal configuration required.

## 🛠️ Installation

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/fjiangAI/ppt-comment-refiner.git
   ```

2. **Navigate to the Project Directory**:

   ```bash
   cd ppt-comment-refiner
   ```

3. **Install Required Dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Usage

This tool uses command-line arguments to specify input and output files as well as the OpenAI API key and other required parameters.

1. **Prepare Your PowerPoint File**: Ensure your `.pptx` file is ready for processing.

2. **Run the Refining and Dubbing Scripts**:

   ### Refining Notes
   First, use the script to extract and enhance your notes:

   ```bash
   python ppt_comment_refiner.py --pptx_path <input.pptx> --output_path <save.pptx> --api_key <api_key> --base_url <base_url>
   ```

   Replace `<input.pptx>` with the path to your original PowerPoint file, `<save.pptx>` with the path to save the refined PowerPoint file, `<api_key>` with your OpenAI API key, and `<base_url>` with the API base URL.

   ### Generate Audio from Notes and Add to PPT
   After enhancing the notes, use the `ppt_dubbing.py` script to generate audio from the refined notes and insert them into the PowerPoint slides:

   ```bash
   python ppt_dubbing.py --pptx_path <input.pptx> --output_pptx_path <output.pptx> --api_key <api_key> --base_url <base_url> --output_notes_file <notes.json> --audio_output_directory <audio_directory> --mp3_directory <audio_directory>
   ```

   Replace `<output.pptx>` with the path to save the updated presentation, and `<audio_directory>` with the directory to save the generated audio files.

## 🤝 Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your improvements.

## 📧 Contact

For questions or support, please open an issue in this repository.

---

*Note: Ensure you have the necessary permissions and API access to use OpenAI's services.*
