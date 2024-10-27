from pdf2image import convert_from_path 
from Converter import Converter, is_sub_identifier, check_create_folder
from datetime import datetime
from collections import OrderedDict
from PIL import Image
import re
import pytesseract 
import pathlib
import os

class Converter_F1(Converter):

    def __init__(self, pdf_file, output_folder, person):
        super().__init__(pdf_file, output_folder, person)
        # Make sure these are identical to output text header with exceptions
        # included
        self.possible_columns = ['Date', 'Job No.', 'ob No.', 'Client', 'From', 'To', 'Service', 'Description', 'Kms', 'Qty', 'UOM', 'Amount', 'Amount $']

        # Add extra re objects
        self.add_re_objects()
        
        # Initalise csv file
        self.extract_text_from_pdf()

    def recreate_columns(self):
        replaces = {"Amount": "Amount $", "ob No.": "Job No."}
        for index, col in enumerate(self.columns):
            if col in replaces:
                self.columns[index] = replaces[col]

    # Driver function to populate self.column_values
    def convert(self):
        self.recreate_columns()

        assert("Amount $" in self.columns)

        column_values = []

        for row in self.rows:
            row_col_values = self.extract_values(row)

            if row_col_values:    
                column_values.append(row_col_values)

                # After a subtotal row, add an empty line
                if row_col_values["Date"] and row_col_values["Service"] == "Total":
                    column_values.append(dict.fromkeys(row_col_values.keys(), ""))

        self.column_values = column_values

        # Correctly order and insert missing col values
        self.columns = []
        for col in self.possible_columns:
            if col in self.column_values[0]:
                self.columns.append(col)

    #  Clean csv by removing unrelevant repetitive info
    def clean_csv(self):
        ''' Overridden from Converter
        '''
        with open(self.csv_file, "r", encoding="ISO-8859-1") as csv_obj:
            cleaned_csv = []
            # start off excluding content until header is encountered
            erase = True

            # content = csv_obj.read().encode(encoding="UTF-8", errors="replace")
            for row in csv_obj.read().split("\n"):
                row = row.strip()
                if row:
                    
                    # Get the actual subtotal value
                    if not self.subtotal_check_val:
                        match = self.regex_name_objs["subtotal_check"].search(row)
                        if match:
                            self.subtotal_check_val = float(match.group(1))
                            self.rows = cleaned_csv
                            return

                    # Stop when encounter "Sub-Total"
                    if "Sub-Total" in row:
                        break

                    # self.header set to None by default and only modified inside this conditional
                    # i.e. self.determine_header must return True before self.header != None
                    if (self.header and row == self.header) or self.determine_header(row):
                        # once the header of the table is encountered, start storing the data
                        erase = False
                        self.header = row
                    elif not erase:
                        # Write data not be erased
                        cleaned_csv.append(self.remove_extra_chars(row))

                    # If there is missing info for header (i.e, atleast 1 None value), try to extract it
                    if not all(self.invoice_info.values()):
                        self.extract_header_info(row)

        self.rows = cleaned_csv

    def extract_values(self, row):
        row_col_values = dict.fromkeys(self.columns, "")
        orig = row
        
        row = self.apply_regexes(row, ["Amount $", "Date"], row_col_values)
        if not row_col_values["Amount $"]:
            return None
        
        row = row.replace("/", "")

        row = self.apply_regexes(row, ["Job No.", "UOM", "Qty", "Kms"], row_col_values)

        row, loc_match = self.apply_loc_regex(row, row_col_values, orig)

        # Get client
        if row_col_values["From"] and row_col_values["To"]:
            # Remove everything before "To" from row
            row_col_values["Client"] = row[:loc_match.span()[0]].strip()
            row = row[loc_match.span()[1]: ]

        # Get service
        row = self.apply_regexes(row, ["Service"], row_col_values)

        # Only thing left should be description
        row_col_values["Description"] = row.strip()

        self.fix_exceptions(row_col_values)

        return row_col_values
    
    def fix_exceptions(self, row_col_values):
        # Fix "2T" Description and Service discrepancy 
        if row_col_values["Service"] in ["27", "2T"]:
            if "2T" in row_col_values["Description"]:
                row_col_values["Service"] = "2T"
            elif "21T HOURLY" in row_col_values["Description"]:
                row_col_values["Service"] = "2T"
                row_col_values["Description"] = row_col_values["Description"].replace("21T HOURLY", "2T HOURLY")
            elif "21 HOURLY" in row_col_values["Description"]:
                row_col_values["Service"] = "2T"
                row_col_values["Description"] = row_col_values["Description"].replace("21 HOURLY", "2T HOURLY")
            

    def apply_loc_regex(self, row, row_col_values, orig):
        reg_obj = self.regex_name_objs["From_To"]
        match = reg_obj.search(row)

        if match:
            row_col_values["From"]  = self.strip_punctuation(match.group(1))
            row_col_values["To"]  = self.strip_punctuation(match.group(2))
        elif row_col_values["Kms"]:
            self.add_new_loc(row, orig)
            self.reset_loc_regex()

            return self.apply_loc_regex(row, row_col_values, orig)

        return row, match
    
    # Get new location from user and update self.locations
    def add_new_loc(self, row, orig):  
        print(orig)
        print(row)
        orig_format = orig.strip().upper()

        entry_from = input(f"\nEnter 'From': ").strip().upper()
        check_sub = f" {entry_from} "
        if entry_from and check_sub in orig_format:
            self.db_m.insert_value(entry_from)

        entry_to = input(f"\nEnter 'To': ").strip().upper()
        check_sub = f" {entry_to} "
        if entry_to and check_sub in orig_format:
            self.db_m.insert_value(entry_to)

    def reset_loc_regex(self):
        self.regex_name_objs["From_To"] = self.compiled_loc_regex()

    # Wrap location regex with re.compile object
    def compiled_loc_regex(self):
        return re.compile(self.get_location_regex())

    # Retrieve locations from db and construct regex
    def get_location_regex(self):
        locations = self.retrieve_locations()
        # Sort locations first based on amount of name components and then total length and join via "|"
        regex = r" *(" + "|".join(locations) + ")"
        # Add space at end to avoid detecting sublocations (like melb in melbourne) matches
        regex += f"{regex} "

        return regex

    def add_re_objects(self):
        # Additional regexes

        # dd/mm/yyy   
        date_regex = r"(?:\d? *\d) */ *(?:\d? *\d) */ *(?:\d *\d *\d *\d|\d *\d)"
        amount_regex = r"\d*\.\d+$"
        qty_regex = amount_regex
        job_no_regex = r"^[, ]*(?![a-z]|\s|,)\d+[., ]* "
        uom_regex = r" [a-z]+ *$"
        kms_regex = r" \d+[, ]*$"
        extra_chars_regex = r"[^a-zA-Z0-9\./\- ]"
        client_regex = r"^ *[A-Z]+ "
        service_regex = r"^ *[a-zA-Z0-9]+[., ]"
        subtotal_check = r"Sub-Total\s*\$\s*(\d+.\d+)$"

        location_regex = self.get_location_regex()
        re_objects = [  ("Qty", qty_regex),
                        ("UOM", uom_regex),
                        ("Kms", kms_regex),
                        ("Amount $", amount_regex),
                        ("Job No.", job_no_regex), 
                        ("Client", client_regex), 
                        ("Date", date_regex),
                        ("Service", service_regex), 
                        ("From_To", location_regex), 
                        ("extra", extra_chars_regex),
                        ("subtotal_check", subtotal_check)
                    ]

        for col, pattern in re_objects:
            # For UOM and Client, case matters
            if col not in ["UOM", "Client"]:
                self.regex_name_objs[col] = re.compile(pattern, flags=re.IGNORECASE)
            else:
                self.regex_name_objs[col] = re.compile(pattern)

    def remove_extra_chars(self, row):
        return re.sub(self.regex_name_objs["extra"], "", row)

    def extract_text_from_pdf(self):
        imgs_folder = f'''imgs_{self.save_name}'''

        pages = convert_from_path(f"{self.pdf_file}", 500) 

        # Creating a text file to write the output 
        outfile = pathlib.Path.cwd().joinpath("out_text.txt")

        if os.path.isdir(imgs_folder):
            self.csv_file = outfile
            return

        check_create_folder(imgs_folder)

        image_counter = 1
        # Iterate through all the pages stored above 
        for page in pages: 
            filename = pathlib.Path.cwd().joinpath(imgs_folder, "page_"+str(image_counter)+".jpg")
            page.save(filename, 'JPEG') 
            image_counter += 1

        # Variable to get count of total number of pages 
        filelimit = image_counter - 1

        f = open(outfile, "a") 
        image_files = []
        # Iterate from 1 to total number of pages 
        for i in range(1, filelimit + 1): 
            filename = pathlib.Path.cwd().joinpath(imgs_folder, "page_"+str(i)+".jpg")
            image_files.append(filename)

            text = str(((pytesseract.image_to_string(Image.open(filename))))) 
            f.write(text) 

        f.close() 

        # Delete the individual images
        for image_file in image_files:
            image_file.unlink()

        # Delete the folder
        pathlib.Path.cwd().joinpath(imgs_folder).rmdir()

        self.csv_file = outfile