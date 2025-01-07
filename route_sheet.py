import logging
from openpyxl import load_workbook
import os
import sys
import json

def load_config(base_dir):
    """
    Load configuration from config.json. If not present, create a default config.
    """
    config_path = os.path.join(base_dir, 'config.json')
    if not os.path.exists(config_path):
        # Create default config
        default_config = {
            "template_path": "assets/RouteSheetTemplateV2.xlsx",
            "output_directory": "assets/generated",
            "log_file": "data/app.log"
        }
        with open(config_path, 'w') as f:
            json.dump(default_config, f, indent=4)
        logging.info("Default config.json created.")
        return default_config
    else:
        with open(config_path, 'r') as f:
            logging.info("config.json loaded.")
            return json.load(f)

def setup_logging(base_dir, log_file):
    """
    Set up logging configuration.
    """
    log_path = os.path.join(base_dir, log_file)
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format='%(asctime)s:%(levelname)s:%(message)s'
    )
    logging.info("Logging is set up.")

def update_route_sheet_from_json(data):
    """
    Update the route sheet based on extracted JSON data.
    
    Parameters:
        data (dict): The JSON data extracted from receipts.
    
    Returns:
        str: The path to the updated route sheet.
    """
    try:
        # Determine the base directory
        if getattr(sys, 'frozen', False):
            # If bundled by PyInstaller
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Load configuration
        config = load_config(base_dir)
        
        # Set up logging
        setup_logging(base_dir, config.get("log_file", "data/app.log"))
        
        # Load the route sheet template using a relative path
        template_path = os.path.join(base_dir, config.get("template_path", "assets/RouteSheetTemplateV2.xlsx"))
        if not os.path.exists(template_path):
            logging.error(f"Template file not found at {template_path}")
            raise FileNotFoundError(f"Template file not found at {template_path}")
        
        workbook = load_workbook(template_path)
        sheet = workbook.active

        logging.info(f"Loaded template from {template_path}")

        # Format dates
        effective_date = data.get("effective_date", "2023-01-01")
        expiration_date = data.get("expiration_date", "2023-12-31")

        # Ensure dates are in MM/DD/YYYY format
        try:
            effective_date_formatted = "/".join(effective_date.split("-")[1:] + [effective_date.split("-")[0]])
            expiration_date_formatted = "/".join(expiration_date.split("-")[1:] + [expiration_date.split("-")[0]])
        except Exception as e:
            logging.error(f"Date formatting error: {e}")
            raise ValueError("Invalid date format in data.")

        # Map extracted fields to cells
        sheet["B4"] = data.get("program", "Unknown Program")
        sheet["C4"] = data.get("council_number", "N/A")
        sheet["D4"] = data.get("district_number", "N/A")
        sheet["G4"] = data.get("local_unit_number", "N/A")
        sheet["H4"] = effective_date_formatted
        sheet["I4"] = data.get("term", "N/A")
        sheet["J4"] = expiration_date_formatted

        logging.info("Mapped basic fields to route sheet.")

        # Write the program type (Troop, Pack, etc.) into E4
        logging.info("Updating E4 with the unit type based on program.")
        program_to_unit_type = {
            "Scouts BSA": "Troop",
            "Cub Scouts": "Pack",
            "Venturing": "Crew",
            "Sea Scouts": "Ship",
            "Exploring": "Post",
            "District": "Non-Unit",
            "Council": "Non-Unit"
        }
        unit_type = program_to_unit_type.get(data.get("program", ""), "Unknown")
        sheet["E4"] = unit_type

        # Map prices to their respective cells
        logging.info("Mapping prices to the correct cells.")
        prices = data.get("prices", {})
        price_mapping = {
            "Charter Renewal": "C8",
            "Youth Registration": "C9",
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
        for field, cell in price_mapping.items():
            if field in prices:
                sheet[cell] = prices.get(field, 0)
                logging.info(f"Set {cell} to {prices.get(field, 0)} for {field}.")

        # Save the updated route sheet with a new name
        district_name = data.get("district_name", "Unknown").replace(" ", "_")
        local_unit_number = data.get("local_unit_number", "Unknown")
        current_date = effective_date_formatted.replace("/", "-")
        
        # Ensure output directory exists
        output_dir = os.path.join(base_dir, config.get("output_directory", "assets/generated"))
        os.makedirs(output_dir, exist_ok=True)
        
        output_filename = f"Route_Sheet_{district_name}_{local_unit_number}_{current_date}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        workbook.save(output_path)
        logging.info(f"Route sheet successfully updated and saved to {output_path}.")
        return output_path

    except Exception as e:
        logging.error(f"Failed to update route sheet: {e}")
        raise e
