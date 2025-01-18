# PowerPoint Audio Automation Tool
# This script automates the extraction of slide notes, audio generation, and inserting the audio into a PowerPoint presentation.

import os
import win32com.client
import json
import time
from pathlib import Path
from openai import OpenAI  # Requires openai library version 1.0+

class AudioGenerator:
    def __init__(self, api_key, base_url):
        self.client = OpenAI(api_key=api_key, base_url=base_url)

    def generate_audio(self, text_list, output_directory):
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        for index, text in enumerate(text_list):
            # Set file path for generated audio
            file_name = f"{output_directory}/{index + 1}.mp3"
            speech_file_path = Path(__file__).parent / file_name

            # Use OpenAI's TTS model to generate audio
            response = self.client.audio.speech.create(
                model="tts-1",  # Model choice
                voice="fable",  # Voice type
                input=text  # Text to be converted to speech
            )

            # Save generated audio to file
            response.stream_to_file(speech_file_path)
            print(file_name)

class PowerPointHandler:
    def __init__(self, pptx_path):
        self.pptx_path = pptx_path
        self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt_app.Visible = True
        self.ppt = None

    def open_presentation(self):
        try:
            pptx_full_path = os.path.abspath(self.pptx_path)
            if not os.path.exists(pptx_full_path):
                raise FileNotFoundError(f"PPT file not found: {pptx_full_path}")

            if not os.access(pptx_full_path, os.R_OK):
                raise PermissionError(f"No permission to read PPT file: {pptx_full_path}")

            self.ppt = self.ppt_app.Presentations.Open(pptx_full_path)
        except Exception as e:
            print(f"Error opening PPT file: {e}")

    def close_presentation(self):
        if self.ppt:
            self.ppt.Close()
        self.ppt_app.Quit()

    def extract_notes(self, output_file):
        if not self.ppt:
            self.open_presentation()

        notes_list = []

        # Extract slide notes
        for slide_idx, slide in enumerate(self.ppt.Slides):
            try:
                # Extract notes from each slide
                notes_text = ""
                for placeholder in slide.NotesPage.Shapes.Placeholders:
                    if placeholder.TextFrame.HasText:
                        notes_text = placeholder.TextFrame.TextRange.Text
                        if notes_text.strip():
                            notes_list.append(notes_text)
                        break
            except Exception as e:
                print(f"Error extracting notes from slide {slide_idx + 1}: {e}")

        # Save notes to JSON file
        with open(output_file, 'w', encoding='utf-8') as notes_file:
            json.dump(notes_list, notes_file, ensure_ascii=False, indent=4)

        print("Notes have been extracted and saved to JSON file.")
        return notes_list

    def insert_audio(self, mp3_directory):
        if not self.ppt:
            self.open_presentation()
        time.sleep(1)  # Delay to ensure each operation completes

        mp3_files = [f for f in os.listdir(mp3_directory) if f.endswith('.mp3')]

        # Insert audio files into each slide
        for slide_idx, slide in enumerate(self.ppt.Slides):
            mp3_file = f"{slide_idx + 1}.mp3"
            if mp3_file in mp3_files:
                mp3_path = os.path.abspath(os.path.join(mp3_directory, mp3_file))

                try:
                    # Insert audio into slide with specified position and size
                    slide.Shapes.AddMediaObject2(mp3_path, LinkToFile=False, SaveWithDocument=True,
                                                 Left=26.5 * 28.35, Top=0.1 * 28.35, Width=1 * 28.35,
                                                 Height=1 * 28.35)
                    time.sleep(0.5)  # Delay for each operation to complete
                except Exception as e:
                    print(f"Error inserting audio into slide {slide_idx + 1}: {e}")

        print("Audio has been inserted into the PPT.")

    def save_presentation(self, output_path):
        if self.ppt:
            self.ppt.SaveAs(os.path.abspath(output_path))
            print("PPT file has been saved.")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='PowerPoint Audio Automation Tool')
    parser.add_argument('--pptx_path', type=str, help='Path to the PowerPoint file')
    parser.add_argument('--output_pptx_path', type=str, help='Path to save the updated PowerPoint file')
    parser.add_argument('--api_key', type=str, help='OpenAI API key')
    parser.add_argument('--base_url', type=str, help='OpenAI API base URL')
    parser.add_argument('--output_notes_file', type=str, help='Path to save extracted notes')
    parser.add_argument('--audio_output_directory', type=str, help='Directory to save generated audio files')
    parser.add_argument('--mp3_directory', type=str, help='Directory containing audio files for insertion')

    args = parser.parse_args()
    audio_api_key = args.api_key
    base_url = args.base_url

    ppt_handler = PowerPointHandler(args.pptx_path)
    audio_generator = AudioGenerator(api_key=audio_api_key, base_url=base_url)

    try:
        # Extract notes and save to file
        notes = ppt_handler.extract_notes(args.output_notes_file)
        print("# Notes extracted and saved.")

        # Generate audio files from notes
        if notes:
            audio_generator.generate_audio(notes, args.audio_output_directory)
            print("# Audio files generated from notes.")

        # Insert audio files into PPT slides
        ppt_handler.insert_audio(args.mp3_directory)
        print("# Audio files inserted into PPT slides.")

        # Save the updated PPTX file
        ppt_handler.save_presentation(args.output_pptx_path)
        print("# Updated PPTX file saved.")
    except Exception as e:
        print(f"Error processing PPT file: {e}")
    finally:
        ppt_handler.close_presentation()

