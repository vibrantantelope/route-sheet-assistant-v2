import pytesseract
from pytesseract import Output
from PIL import Image, ImageEnhance, ImageFilter
from pdf2image import convert_from_path
import json
import os
import sys
from datetime import datetime, timedelta
import re
import logging

# Ensure the data directory exists
os.makedirs("data", exist_ok=True)

# Set up logging
logging.basicConfig(
    filename="data/extract_receipt_debug.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def load_config(base_dir):
    """
    Load the application's configuration from config.json.
    """
    config_path = os.path.join(base_dir, "config.json")
    if not os.path.exists(config_path):
        # Create a default config if not present
        default_config = {
            "tesseract_path": "tesseract/Tesseract-OCR/tesseract.exe",
            "poppler_path": "poppler/poppler-24.08.0/Library/bin"
        }
        with open(config_path, "w") as f:
            json.dump(default_config, f, indent=4)
        logging.info("Default config.json created.")
        return default_config
    else:
        with open(config_path, "r") as f:
            logging.info("config.json loaded.")
            return json.load(f)

# Determine the base directory
if getattr(sys, "frozen", False):
    # If bundled by PyInstaller
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Load configuration
config = load_config(base_dir)

# Set the Tesseract path from config.json
tesseract_path = os.path.join(base_dir, config.get("tesseract_path", "tesseract/Tesseract-OCR/tesseract.exe"))
pytesseract.pytesseract.tesseract_cmd = tesseract_path
print(f"Tesseract path being used: {pytesseract.pytesseract.tesseract_cmd}")


# Set the Poppler path from config.json
poppler_path = os.path.join(base_dir, config.get("poppler_path", "poppler/poppler-24.08.0/Library/bin"))

def preprocess_image(image):
    """
    Preprocess the image to improve OCR accuracy.
    """
    logging.info("Preprocessing image for OCR.")
    # Convert to grayscale
    image = image.convert("L")
    # Resize the image for better OCR
    image = image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)
    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(3)
    # Apply sharpening filter
    image = image.filter(ImageFilter.SHARPEN)
    logging.info("Image preprocessing complete.")
    return image

def convert_pdf_to_image(pdf_path):
    """
    Convert a PDF file to an image using the bundled Poppler.
    """
    try:
        logging.info(f"Converting PDF to image: {pdf_path}")
        images = convert_from_path(pdf_path, dpi=400, poppler_path=poppler_path)
        logging.info("PDF conversion complete.")
        return images[0]  # Use the first page
    except Exception as e:
        logging.error(f"Error converting PDF to image: {e}")
        raise RuntimeError(f"Failed to convert PDF to image: {e}")

def extract_receipt_text_to_json(receipt_path):
    """
    Extract text from receipt (PDF or image) using OCR and save it as a JSON file.
    """
    try:
        logging.info(f"Processing receipt: {receipt_path}")

        # Convert PDF to image if necessary
        if receipt_path.lower().endswith(".pdf"):
            image = convert_pdf_to_image(receipt_path)
        else:
            # Load image directly
            logging.info("Loading image file.")
            image = Image.open(receipt_path)

        # Preprocess the image
        image = preprocess_image(image)

        # Perform OCR on the image
        text = pytesseract.image_to_string(image, lang='eng', config='--psm 4')

        # Debugging: Save raw OCR output for review
        with open("data/raw_ocr_output.txt", "w") as f:
            f.write(text)
        logging.debug("OCR output saved to raw_ocr_output.txt.")

        # Initialize data dictionary
        data = {
            "council_number": "456",
            "effective_date": datetime.now().strftime("%Y-%m-01"),
            "term": "12 months",
        }

        # Extract district, unit, program type, and price information
        district_map = {
            "Calumet": 1, "Aguila": 2, "Prairie Dunes": 3, "Thunderbird": 4, "Checaugau": 5,
            "Iron Horse": 6, "Tri-Star": 7, "Five Creeks": 9, "Tall Grass": 11, "Trailblazer": 12
        }

        price_fields = {
            "Youth Registration": "C9",
            "Charter Renewal": "C8",
            "Youth SL Subscription": "C10",
            "Youth Transfer": "C11",
            "Adult Registration": "C12",
            "Multiple/Position Change": "C13",
            "Adult Transfer": "C14",
            "Adult SL Subscription": "C15",
            "Youth Exploring": "C16",
            "Adult Exploring": "C17",
            "Program Fee": "C18",
        }

        prices = {field: 0 for field in price_fields.keys()}

        # Process each line for data
        logging.info("Processing OCR lines for data extraction.")
        lines = text.splitlines()
        for line in lines:
            logging.debug(f"Processing line: {line}")
            line = line.strip()

            # Extract district
            for district, number in district_map.items():
                if district.lower() in line.lower():
                    data["district_name"] = district
                    data["district_number"] = number
                    logging.info(f"Matched district: {district} ({number})")

            # Extract local unit number
            if "Troop" in line or "Pack" in line or "Crew" in line or "Ship" in line or "Post" in line:
                match = re.search(r"(\d+)", line)
                if match:
                    data["local_unit_number"] = match.group(1)
                    logging.info(f"Matched local unit number: {data['local_unit_number']}")

            # Check for Charter Renewal
            # Check for Charter Renewal or Unit Charter
            if "Charter Renewal" in line or "Unit Charter" in line:
                match = re.search(r"(\d+)", line)
                if match:
                    quantity = int(match.group(1))
                    if "Unit Charter" in line:
                        # Adjust quantity to 1 if it was incorrectly set to 100
                        quantity = 1 if quantity == 100 else quantity
                    prices["Charter Renewal"] += quantity
                    data["charter_renewal"] = quantity
                    logging.info(f"Charter Renewal captured: {quantity}")



            # Extract program type
            if "Scouts BSA" in line or "Troop" in line:
                data["program"] = "Scouts BSA"
            elif "Cub Scouts" in line or "Pack" in line:
                data["program"] = "Cub Scouts"
            elif "Venturing" in line or "Crew" in line:
                data["program"] = "Venturing"
            elif "Sea Scouts" in line or "Ship" in line:
                data["program"] = "Sea Scouts"
            elif "Exploring" in line or "Post" in line:
                data["program"] = "Exploring"

            # Extract price information
            match = re.search(r"(\d+)\s+(Youth BL|Youth Renewal|Adult Renewal|Adult New|Youth Program Fee|Adult Program Fee)", line, re.IGNORECASE)
            if match:
                count = int(match.group(1))
                label = match.group(2).strip().lower()
                logging.info(f"Matched price line: {label} ({count})")

                if "youth bl" in label:
                    prices["Youth SL Subscription"] += count
                elif "youth renewal" in label or "youth new" in label:
                    prices["Youth Registration"] += count
                elif "adult renewal" in label or "adult new" in label:
                    prices["Adult Registration"] += count
                elif "program fee" in label:
                    prices["Program Fee"] += count

        # Add prices to data
        data["prices"] = prices

        # Infer expiration date
        effective_date = datetime.strptime(data["effective_date"], "%Y-%m-%d")
        expiration_date = effective_date + timedelta(days=365) - timedelta(days=1)
        data["expiration_date"] = expiration_date.strftime("%Y-%m-%d")

        # Save data to JSON
        json_path = "data/receipt_data.json"
        with open(json_path, "w") as json_file:
            json.dump(data, json_file, indent=4)
        logging.info("Data successfully extracted and saved to JSON.")

        return data

    except Exception as e:
        logging.error(f"Error processing receipt: {e}")
        raise RuntimeError(f"Failed to process receipt: {e}")
