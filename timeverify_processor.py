import pytesseract
from docx import Document
from PIL import Image, ImageEnhance
import pandas as pd
import io
import os
import re
from datetime import datetime

class TimesheetProcessor:
    def __init__(self):
        # Set tesseract path for Windows
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        self.timesheet_data = []

    def extract_images_from_word_file(self, word_file_path):
        """Extract images from Word document - YOUR EXISTING METHOD"""
        images = []
        try:
            doc = Document(word_file_path)
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    try:
                        image_data = rel.target_part.blob
                        image = Image.open(io.BytesIO(image_data))
                        if image.mode != 'RGB':
                            image = image.convert('RGB')
                        images.append(image)
                    except:
                        continue
        except Exception as e:
            print(f"Error with document: {e}")
        return images

    def extract_name_from_filename(self, filename):
        """Extract name from filename - YOUR EXISTING METHOD"""
        name = filename.replace('.docx', '').replace('.doc', '')
        name = re.sub(r'-\d{4}', '', name)  # Remove year
        name = re.sub(r'[_-]', ' ', name)   # Replace underscores/dashes with spaces
        name = name.strip()
        return name if name else "Unknown"

    def extract_text_from_image(self, image):
        """Extract text using multiple OCR methods - YOUR EXISTING METHOD"""
        all_text = []
        
        # Method 1: Standard OCR
        try:
            config = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'
            text = pytesseract.image_to_string(image, config=config)
            all_text.append(text)
        except:
            pass
        
        # Method 2: Enhanced contrast
        try:
            enhancer = ImageEnhance.Contrast(image)
            enhanced = enhancer.enhance(2.0)
            text = pytesseract.image_to_string(enhanced, config=config)
            all_text.append(text)
        except:
            pass
        
        # Method 3: Different PSM mode
        try:
            config = r'--oem 3 --psm 4'
            text = pytesseract.image_to_string(image, config=config)
            all_text.append(text)
        except:
            pass
        
        # Combine all text
        return '\n'.join(all_text)

    def parse_timesheet_entries(self, text, employee_name):
        """Parse ALL date and hour entries from text - YOUR EXISTING METHOD"""
        entries = []
        if not text:
            return entries
            
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        for line in lines:
            # Look for date patterns
            date_patterns = [
                r'(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2}/\d{1,2}/\d{4})',
                r'(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{1,2}/\d{1,2}/\d{4})',
                r'(\d{1,2}/\d{1,2}/\d{4})',
                r'(\d{1,2}-\d{1,2}-\d{4})',
            ]
            
            date_match = None
            for pattern in date_patterns:
                match = re.search(pattern, line)
                if match:
                    date_match = match
                    break
            
            if date_match:
                date = date_match.group(1)
                
                # Look for hours in the same line
                hours_patterns = [
                    r'(\d{1,2}(?:\.\d{1,2})?)\s+(?:Cost|Installation|Store|Enterprise|Product|DACI|Post|Item|Hours?|Hrs?)',
                    r'Product\s*/\s*[A-Za-z\s]+\s*/\s*(\d{1,2}(?:\.\d{1,2})?)',
                    r'\(\s*(\d{1,2}(?:\.\d{1,2})?)\s*\)',
                    r'(?<=\s)(\d{1,2}(?:\.\d{1,2})?)(?=\s+)',
                    r'/\s*(\d{1,2}(?:\.\d{1,2})?)\s*',
                    r'(?<=\s)(\d{1,2}(?:\.\d{1,2})?)\.?0?(?=\s|$)',
                    r'(?<=\s)(\d{1,2}(?:\.\d{1,2})?)$',
                    r'(\d{1,2}(?:\.\d{1,2})?)\s*h(?:r|rs?|ours?)?',
                    r'(?:^|\s)(\d{1,2}(?:\.\d{1,2})?)(?=\s|$)'
                ]
                
                hours_match = None
                for pattern in hours_patterns:
                    match = re.search(pattern, line)
                    if match:
                        hours_match = match
                        break
                
                if hours_match:
                    try:
                        hours = float(hours_match.group(1))
                        if 0 <= hours <= 24:  # Validate hours
                            # Parse and format date
                            try:
                                if '/' in date:
                                    parsed_date = pd.to_datetime(date, format='%m/%d/%Y')
                                else:
                                    parsed_date = pd.to_datetime(date, format='%m-%d-%Y')
                                
                                formatted_date = parsed_date.strftime('%m/%d/%Y')
                                
                                entry = {
                                    'Name': employee_name,
                                    'Date': formatted_date,
                                    'Hours': hours
                                }
                                entries.append(entry)
                            except:
                                # Add with CHECK if date parsing fails
                                entry = {
                                    'Name': employee_name,
                                    'Date': date,
                                    'Hours': hours
                                }
                                entries.append(entry)
                    except:
                        pass
                else:
                    # Date found but no hours
                    entry = {
                        'Name': employee_name,
                        'Date': date,
                        'Hours': "CHECK"
                    }
                    entries.append(entry)
        
        return entries

    def simulate_ibm_system_check(self, consultant_name):
        """Simulate IBM system lookup for demo purposes"""
        demo_data = {
            'john smith': 38,   # Screenshot might show 40, system shows 38
            'jane doe': 40,     # Perfect match
            'mike johnson': 35, # Perfect match
            'sarah wilson': 40, # Screenshot might show 42, system shows 40
            'alex chen': 38     # Slight discrepancy
        }
        
        name_key = consultant_name.lower().strip()
        return demo_data.get(name_key, 40)  # Default to 40 if not found