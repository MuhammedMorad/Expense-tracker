import speech_recognition as sr
from pydub import AudioSegment
import os
import re

class SpeechHandler:
    def __init__(self):
        self.recognizer = sr.Recognizer()

    def convert_to_wav(self, input_path):
        """Converts input audio to WAV format required by SpeechRecognition."""
        try:
            output_path = input_path + ".wav"
            # Explicitly set ffmpeg location if needed, but it should be in PATH
            audio = AudioSegment.from_file(input_path)
            audio.export(output_path, format="wav")
            return output_path
        except Exception as e:
            print(f"Error converting audio: {e}")
            return None

    def transcribe(self, audio_path, language="ar-EG"):
        """Transcribes audio file to text."""
        wav_path = self.convert_to_wav(audio_path)
        if not wav_path:
            return None

        try:
            with sr.AudioFile(wav_path) as source:
                audio_data = self.recognizer.record(source)
                text = self.recognizer.recognize_google(audio_data, language=language)
                return text
        except sr.UnknownValueError:
            return "" # Could not understand audio
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            return None
        finally:
            # Cleanup wav file
            if wav_path and os.path.exists(wav_path):
                os.remove(wav_path)

    def parse_text(self, text):
        """
        Extracts amount, description, year, and expense type from text.
        Heuristics:
        - Amount: First number sequence.
        - Year: 4-digit number starting with 20 (e.g., 2023, 2024, 2025).
        - Type: Keywords 'أساسي' or 'فرعي'.
        - Description: The rest.
        """
        if not text:
            return None, None, None, None

        # 1. Extract Year
        year = None
        year_match = re.search(r'\b(20\d{2})\b', text)
        if year_match:
            year = year_match.group(1)
            text = text.replace(year, "", 1) # Remove year from text

        # 2. Extract Expense Type
        expense_type = None
        if re.search(r'(أساسي|اساسي)', text):
            expense_type = 'Essential'
            text = re.sub(r'(أساسي|اساسي)', '', text)
        elif re.search(r'(فرعي|جانبي)', text):
            expense_type = 'Side'
            text = re.sub(r'(فرعي|جانبي)', '', text)

        # 3. Extract Amount
        amount = 0
        amount_match = re.search(r'(\d+(\.\d+)?)', text)
        if amount_match:
            amount_str = amount_match.group(1)
            try:
                amount = float(amount_str)
            except ValueError:
                amount = 0
            text = text.replace(amount_str, "", 1) # Remove amount from text

        # 4. Cleanup Description
        stop_words = [
            'صرفت', 'دفعت', 'اشتريت', 
            'جنيه', 'ريال', 'دولار', 'ليرة', 
            'في', 'على', 'ب', 'من', 'سنة', 'عام',
            'مصروف', 'نوع'
        ]
        
        desc_words = text.split()
        filtered_words = [w for w in desc_words if w.strip() not in stop_words]
        description = " ".join(filtered_words)

        return amount, description.strip(), year, expense_type
