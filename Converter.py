from DBManagerLocation import DBManagerLocation
from collections import OrderedDict
from datetime import datetime
import pathlib
import tabula
import csv
import re
import os

# Check if a folder exists in the current dir, if not create
def check_create_folder(folder_name):
    if not os.path.isdir(folder_name):
        os.makedirs(folder_name)
    
def is_sub_identifier(row):
    identifier = "−"
    if row[0] == identifier:
        return row == identifier * len(row)
    return False

# Check if "Amount" if the only key with a valid value
def is_subtotal(row_dict):
    return row_dict["Amount"] == "".join(row_dict.values())

class Converter:

    def __init__(self, pdf_file, output_folder, person):
        self.pdf_file = pdf_file
        self.dir = os.path.dirname(self.pdf_file)
        self.person = person
        self.filename, _ = os.path.splitext(os.path.basename(pdf_file))
        self.save_name = f'''{datetime.now().strftime("%d-%m-%Y %H_%M_%S")} {self.filename}'''

        print(self.save_name)

        self.output_folder = pathlib.Path.cwd().joinpath(output_folder)
        self.logs_folder = pathlib.Path.cwd().joinpath(output_folder , "logs")

        self.subtotal_seperator = "--------------------"
        self.after_date_toll_text = "TL"

        self.db_m = DBManagerLocation("locations.db")
        self.__regex_name_objs = self.get_re_objects()

        self.csv_file = None

        self.__possible_columns = ['Date', 'Job No', 'Client', 'Reference', 'From', 'To', 'Rate', 'Amount', 'Cash']
        self.columns = None
        self.column_values = None

        self.__subtotal_check_val = None

        # Column header row in pdf 
        self.header = None
        # Invoice info
        self.__invoice_info = {"Invoice Date:": None, "Period Ending:": None}
        # Good and services info
        self.__goods_info = []
        self.rows = None

    # Driver function to populate self.column_values
    def convert(self):
        assert("Amount" in self.columns or "Amount $" in self.columns)

        column_values = []
        last_after_date_toll_index = -1
        for row in self.rows:

            if is_sub_identifier(row):
                # For row before subtotal amount, add the default toll text before amount
                row_col_values = dict.fromkeys(self.columns, "")
                row_col_values[self.get_col_bef("Amount")] = self.after_date_toll_text
                column_values.append(row_col_values)

                # For row before subtotal-amount, add seperator value to column before amount
                row_col_values = dict.fromkeys(self.columns, "")
                row_col_values[self.get_col_bef("Amount")] = self.get_subtotal_seperator()
            else:
                row_col_values = self.extract_values(row)

            column_values.append(row_col_values)

            # If row is subtotal row, add an 'empty' line after it
            if is_subtotal(row_col_values):
                column_values.append(dict.fromkeys(row_col_values.keys(), ""))

        # The last after date toll index is put before the final total value which is unnecessary
        last_after_date_toll_index = -1
        for i in range(len(column_values) - 1, 0, -1):
            if column_values[i][self.get_col_bef("Amount")] == self.after_date_toll_text:
                last_after_date_toll_index = i
                break

        column_values.pop(last_after_date_toll_index)

        self.column_values = column_values


        # Treat goods info and convert to (desc, val) dict
        self.extract_goods_info()

    # Convert file to csv and store in (created) self.output_folder 
    def convert_to_csv(self):
        self.csv_file = pathlib.Path.cwd().joinpath(self.dir, self.filename + ".csv")

        if not os.path.isfile(self.csv_file):
            tabula.convert_into(self.pdf_file, str(self.csv_file), output_format="csv", stream=False, lattice=True, pages="all")

    # Clean csv by removing unrelevant repetitive info
    def clean_csv(self):
        self.convert_to_csv()

        with open(self.csv_file, "r") as csv_obj:
            cleaned_csv = []
            first_row = None
            erase = False
            record_goods = False

            reader = csv.reader(csv_obj)
            for elem in reader:
                # The first value in each element stores the data
                for row in elem[0].split("\n"):
                    row = row.strip()

                    # Record the first row to mark the start of areas to be deleted
                    if not first_row:
                        first_row = row

                    # Get the actual subtotal value
                    if not self.subtotal_check_val:
                        match = self.regex_name_objs["subtotal_check"].search(row)
                        if match:
                            self.subtotal_check_val = float(match.group(1))

                    # Once the first row is encountered, start erasing
                    if row == first_row:
                        erase = True

                    # self.header set to None by default and only modified inside this conditional
                    # i.e. self.determine_header must return True before self.header != None
                    if (self.header and row == self.header) or self.determine_header(row):
                        # once the header of the table is encountered, start storing the data
                        erase = False
                        self.header = row
                    elif not erase:
                        # Write data not be erased
                        cleaned_csv.append(row)

                    # If there is missing info for header (i.e, atleast 1 None value), try to extract it
                    if not all(self.__invoice_info.values()):
                        self.extract_header_info(row)

                    # If no goods_info, check if it can start recording 
                    if not self.__goods_info:
                        record_goods = self.start_goods_info(row)
                    elif is_sub_identifier(row):
                        # If there is goods_info recorded and a separator is encountered, stop recording goods data
                        record_goods = False

                    if record_goods:
                        # if allowed, start recording goods info
                        self.__goods_info.append(row)

            self.rows = cleaned_csv

    def extract_header_info(self, row):
        for info_needed in self.__invoice_info.keys():
            if not self.__invoice_info[info_needed]:
                self.apply_regex(info_needed, self.regex_name_objs[info_needed], row, self.__invoice_info, default_val=None, default_group=1)

    def start_goods_info(self, row):
        match = re.search(self.regex_name_objs["Goods"], row)
        return bool(match)
 
    def extract_goods_info(self):
        goods_info = {}
        # Remove the "Goods and services supplied;" header included due to regex match
        self.__goods_info = self.__goods_info[1:]

        for goods_row in self.__goods_info:
            split_row = [comp for comp in goods_row.split(" ") if comp]
            # The last value is amount, and rest description of good
            desc, val = " ".join(split_row[:-1]), split_row[-1]

            # Cast to floats if possible
            try:
                val = float(val)
            except ValueError:
                pass

            goods_info[desc] = val

        # Reassign treated dict info struct
        self.__goods_info = goods_info

    # Return if row is the column header of pdf
    # if yes, set self.columns
    def determine_header(self, row):
        # If the column header has already been found or the first column doesn't match
        if self.columns and row.split(" ")[0] != self.possible_columns[0]:
            return False

        header_cols = []
        # Check each possible col 
        for col in self.possible_columns:
            if row and col in row:
                header_cols.append(col)
                # Remove from row the columns that match
                row = row.replace(col, "").strip()

        # If row still has elements, row has non-column values i.e., it's not a column header
        self.columns = header_cols if not row else self.columns
        return not row

    # Extract columns before and incl. 'To'
    def extract_cols(self, row, cols, row_col_values):
        assert(self.header)
        row = row.strip()

        for col in cols:
            assert(self.columns.index(col) <= self.columns.index("To"))

            if row:
                col_index = self.header.find(col)
                next_col_index = self.header.find(self.get_col_after(col)) - col_index
                
                val = row[:next_col_index].strip()
                # Erase extracted col data from row 
                row = row[next_col_index:]

                row_col_values[col] = val.strip()

    def extract_values(self, row):
        row_col_values = dict.fromkeys(self.columns, "")
        
        row = self.apply_regexes(row, ("Amount", "Rate"), row_col_values)
        row = self.apply_regexes(row, ("Date", "Job No"), row_col_values)

        cols = [col for col in self.columns if col not in ["Amount", "Rate", "Date", "Job No"]]
        self.extract_cols(row, self.get_cols_before("To", cols), row_col_values)

        return row_col_values

    # cols is a tuple of relevant column names to apply regexes and extract their values
    def apply_regexes(self, row, cols, row_col_values):

        for col in cols:
            row = self.apply_regex(col, self.regex_name_objs[col], row, row_col_values)
            # Prepend ' ' to row if needed to allow 'rate' regex to match
            row = f" {row}" if row and row[0] not in [",", " "] else row

        return row

    # Apply a single regex compiled object to row and store extracted values in row_col_values
    def apply_regex(self, column_name, reg_obj, row, row_col_values, default_val="", default_group=0):
        match = reg_obj.search(row)

        if match:
            row_col_values[column_name]  = self.strip_punctuation(match.group(default_group))
            # Remove stripped match from row string
            row = reg_obj.sub("", row, count=1)
        else:
            row_col_values[column_name]  = default_val

        return row

    # Compile and return regex patterns as regex objects as dict
    def get_re_objects(self):

        # dd Month_Name yyyy
        date_regex = r"^[,\s]*(\s*\d\d?)\s*(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?){1}(\s*\d{2,4})"
        # dd/mm/yyy   
        amount_regex = r"\d*\.\d+$"
        job_no_regex = r"^[, ]*(?![a-z]|\s|,)\d+ "
        rate_regex = r"(?<=,|\s)[a-z0-9]+(?=[,\s]*$)"
        invoice_date = r"invoice\s*date\s*:\s*(.+)$"
        period_ending = r"period\s*ending\s*:\s*(.+)$"
        goods_regex = r"^Goods and services supplied;$"
        punc_strip = r"^[,. ]+|[,. ]+$"
        subtotal_check = r"total freight charges\s*(\d*\.*\d+)"

        re_objects = [ 
                       ("Date", date_regex),
                       ("Amount", amount_regex),
                       ("Job No", job_no_regex), 
                       ("Rate", rate_regex), 
                       ("Invoice Date:", invoice_date),
                       ("Period Ending:", period_ending),
                       ("Goods", goods_regex),
                       ("punc_strip", punc_strip),
                       ("subtotal_check", subtotal_check)
                    ]
        # Map regex expressions to compiled objects except 'Goods'
        re_objects = [(col, re.compile(pattern, flags=re.IGNORECASE)) for col, pattern in re_objects]

        # Order of applying regex exps matters therefore OrderedDict
        re_objects = OrderedDict(re_objects)

        return re_objects

    # Return locations from db sorted first by number of name components and then by total length
    def retrieve_locations(self):
        locations = self.db_m.get_locations()
        locations = sorted(locations, key=lambda loc: (len(loc.split(" ")), len(loc)), reverse=True)

        return locations

    def strip_punctuation(self, content):
        return re.sub(self.regex_name_objs["punc_strip"], "", content)
        # return re.sub("^[, \\t]+|[, \\t]+$", "", content)

    # GETTERS --------------------------------------------------------------------

    @property
    def invoice_info(self):
        return self.__invoice_info
    
    @invoice_info.setter
    def invoice_info(self, value):
        self.__invoice_info = value

    @property
    def goods_info(self):
        return self.__goods_info


    @property
    def possible_columns(self):
        return self.__possible_columns
    
    @possible_columns.setter
    def possible_columns(self, value):
        self.__possible_columns = value

    @property
    def regex_name_objs(self):
        return self.__regex_name_objs

    def get_col_bef(self, col):
        index_col_bef = self.columns.index(col) - 1
        if index_col_bef >= 0:
            return self.columns[index_col_bef]
        else:
            # If no column before col
            return col

    def get_col_after(self, col):
        index_col_after = self.columns.index(col) + 1
        if index_col_after < len(self.columns):
            return self.columns[index_col_after]
        else:
            # If no column after col
            return col

    def get_cols_before(self, col, cols):
        col_index = cols.index(col)
        return [cols[i] for i in range(len(cols)) if i <= col_index]

    def get_subtotal_seperator(self):
        return self.subtotal_seperator

    def get_column_values(self):
        return self.column_values

    def get_columns(self):
        return self.columns
    
    def get_output_folder(self):
        return self.output_folder
    
    def get_filename(self):
        return self.filename

    @property
    def subtotal_check_val(self):
        return self.__subtotal_check_val

    @subtotal_check_val.setter
    def subtotal_check_val(self, value):
        self.__subtotal_check_val = value