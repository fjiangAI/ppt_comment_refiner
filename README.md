
# 🎨 PPT Comment Refiner

Enhance your PowerPoint presentation notes effortlessly using OpenAI's GPT model.

## 📋 Overview

PPT Comment Refiner is a Python tool designed to refine and enhance the notes in your PowerPoint presentations. By leveraging OpenAI's GPT model, it improves the fluency, coherence, and overall quality of your presentation notes, making them more suitable for academic and professional settings.

## ✨ Features

- **Automated Note Enhancement**: Utilizes OpenAI's GPT model to refine presentation notes for better clarity and eloquence.
- **Seamless Integration**: Processes existing PowerPoint files and outputs refined versions without manual intervention.
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

This tool uses command-line arguments to specify input and output files as well as the OpenAI API key.

1. **Prepare Your PowerPoint File**: Ensure your `.pptx` file is ready for processing.

2. **Run the Script**:

   ```bash
   python ppt_comment_refiner.py <input.pptx> <output.pptx> <api_key>
   ```

   Replace `<input.pptx>` with the path to your original PowerPoint file, `<output.pptx>` with the path to save the refined presentation, and `<api_key>` with your OpenAI API key.

3. **Output**: The refined PowerPoint file will be saved to the specified output path.

## 🤝 Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your improvements.

## 📧 Contact

For questions or support, please open an issue in this repository.

---

*Note: Ensure you have the necessary permissions and API access to use OpenAI's services.*
