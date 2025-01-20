import os
import subprocess

def main():
    print("Welcome to the Excel Conversion and Processing Tool!")

    # Prompt the user to enter the source Excel file path
    source_file = input("Please enter the path to the source Excel file: ").strip()

    # Check if the source file exists
    if not os.path.isfile(source_file):
        print("Error: The specified file does not exist. Please try again.")
        return

    # Run the first script: `csv_generator.py` to convert Excel to CSVs
    print("\nStep 1: Converting Excel sheets into CSV files...")
    try:
        subprocess.run(["python", "csv_generator.py", source_file], check=True)
        print("Conversion completed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error during CSV generation: {e}")
        return

    # Assume the folder `converted_csvs` is created by the first script
    converted_csv_folder = "converted_csvs"
    if not os.path.isdir(converted_csv_folder):
        print(f"Error: The folder '{converted_csv_folder}' was not found after running the first script.")
        return

    # Run the second script: `sheet_generator_v2.py` to process the CSVs and create a final Excel file
    print("\nStep 2: Compiling CSV files into a new Excel file...")
    try:
        subprocess.run(["python", "sheet_generator_v2.py", converted_csv_folder], check=True)
        print("Final Excel file generated successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error during Excel file generation: {e}")
        return

    print("\nProcess completed successfully. Your final Excel file is ready!")

if __name__ == "__main__":
    main()
