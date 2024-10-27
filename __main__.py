from DBManagerLocation import DBManagerLocation
from Converter_F1 import Converter_F1
from ExcelWriter import ExcelWriter
from ExcelWriter_F1 import ExcelWriter_F1
from Converter import Converter
from datetime import date
import os
import pathlib

def get_account(file):
    file = str(file).upper()

    # Gather account names from Acc in format_info 
    identities = [acc.stem for acc in pathlib.Path.cwd().parents[0].joinpath("Acc").glob("*.txt")]

    account = input(f"Enter name of account out of {identities}: ").strip().upper()

    if account not in identities:
        print(f'''The filename couldn't be associated with any of the registered accounts.''')
        return get_account(file)

    return account

def multifile():
    os.system("clear")
    folders = { "F1": pathlib.Path.cwd().parents[0].joinpath("F1 Do"), 
                "F2": pathlib.Path.cwd().parents[0].joinpath("F2 Do")
              }

    for format in folders:
        folder_path = folders[format]
        if os.path.isdir(folder_path):

            print(f"\nFiles inside {folder_path}\n")
            output_folder = pathlib.Path.cwd().parents[0].joinpath("OUTPUT", date.today().strftime("%d-%m-%Y"))

            for file in folder_path.glob("*.pdf"):
                _filename, ext = os.path.splitext(file)

                file = pathlib.Path.cwd().joinpath(folder_path, file)
                print(file.name)

                if ext != ".pdf":
                    print("File is not PDF.")
                    continue

                # Write to excel
                person = get_account(file)

                percent_dis = None
                while not percent_dis:
                    try:
                        percent_dis = float(input(f"Enter percent discount for {file} (like 70 or 70%): ").replace("%", ""))
                    except ValueError:
                        print("Please make sure the percent discount is a number!")

                # Get Gst Input
                incl_gst = input("Are prices inclusive of GST [Y/N]: ").lower()
                if not (incl_gst == "y" or incl_gst == "n"):
                    print("Enter either Y or N")
                    incl_gst = input("Are prices inclusive of GST [Y/N]: ").lower()

                # Get Vehical Type Input
                veh_type = input("Is the invoice for Heavy(H) or Light(L) vehicle: ").lower()
                if not (veh_type == "h" or veh_type == "l"):
                    print("Enter either H or L")
                    veh_type = input("Is the invoice for Heavy(H) or Light(L) vehicle: ").lower()

                if format == "F1":
                    converter = Converter_F1(file, output_folder, person)
                    converter.clean_csv()
                    converter.convert()
                    writer = ExcelWriter_F1(converter, percent_dis, incl_gst)
                elif format == "F2":
                    converter = Converter(file, output_folder, person)
                    converter.clean_csv()
                    converter.convert()
                    writer = ExcelWriter(converter, percent_dis, incl_gst)
                else:
                    print("Invalid input folder. Move file to appropriate folder and try again.")
                    return

                writer.config()
                writer.write_wb()
                                
                try:
                    writer.subtotal_test()
                except:
                    print("Something went wrong!")
                            
                writer.write_summary()
                writer.save()

                input("\nPress enter to continue to next one...")
                os.system("clear")

multifile()

