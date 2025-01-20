import pandas as pd
import os
import argparse

def convert_excel_to_csv(excel_file, output_folder):
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel_data = pd.ExcelFile(excel_file)

    for sheet_name in excel_data.sheet_names:
        df = excel_data.parse(sheet_name)

        csv_file = os.path.join(output_folder,f"{sheet_name}.csv")

        df.to_csv(csv_file, index=False)
        print(f"converted sheet '{sheet_name}' to '{csv_file}'")

if __name__ ==  "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel sheets to CSV files")
    parser.add_argument("excel_file", help="Path to the excel file")
    parser.add_argument(
        "--output_folder",
        default = "converted_csvs",
        help="Output folder for csv files (default: 'converted_csvs')"
    )

    args = parser.parse_args()
    convert_excel_to_csv(args.excel_file, args.output_folder)