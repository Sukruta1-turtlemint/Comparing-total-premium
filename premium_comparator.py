import os
import pandas as pd
import glob
from openpyxl import load_workbook
import pyxlsb
import xlrd
from pyxlsb import open_workbook

class PremiumComparator:
    def __init__(self, base_folder, s3_premium_file, output_folder):
        """
        Initializes the PremiumComparator class with file paths.
        """
        self.base_folder = base_folder
        self.s3_premium_file = s3_premium_file
        self.output_folder = output_folder
        self.given_premium_file = os.path.join(output_folder, "Given_Premium_2022-23.xlsx")
        self.refined_premium_file = os.path.join(output_folder, "Refined_Given_Premium_2022-23.xlsx")
        self.comparison_file = os.path.join(output_folder, "Comparison_2022-23.xlsx")

        # Create output folder if not exists
        os.makedirs(output_folder, exist_ok=True)

    def extract_total_premium(self, file_path, sheet_name, file_type):
        """
        Extracts the sum of the 'Total premium' column from a given sheet, dynamically detecting the header row.
        """
        try:
            if file_type == ".xlsx":
                temp_df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine="openpyxl",header=None)
            elif file_type == ".xls":
                temp_df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine="xlrd",header=None)
            elif file_type == ".xlsb":
                with pyxlsb.open_workbook(file_path) as wb:
                    with wb.get_sheet(sheet_name) as sheet:
                    # Read only the first 5 rows
                        data = [ [cell.v for cell in row] for _, row in zip(range(5), sheet.rows()) ]  
                        # Convert to DataFrame without assigning headers
                        temp_df = pd.DataFrame(data)
            else:
                return 0
            print("temp_df:", temp_df)

            correct_header_row = None
            for i, row in temp_df.iterrows():
                # Remove NaNs, convert to string, strip spaces, and convert to lowercase
                row_values = [str(val).strip().lower() for val in row.dropna().tolist()]
                print(f"Row {i}: {row_values}")  # Debugging output

                # Check for "total premium" in a case-insensitive manner
                if "total premium" in row_values:
                    correct_header_row = i
                    break

            if correct_header_row is not None:
                print(f"✅ Detected Header Row: {correct_header_row}")
            else:
                print("⚠️ Warning: Could not detect header row!")


            # Read the full sheet with the detected header row
            if file_type == ".xlsx":
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=correct_header_row, engine="openpyxl")
            elif file_type == ".xls":
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=correct_header_row, engine="xlrd")
            elif file_type == ".xlsb":
                # Load the full sheet using pyxlsb
                with pyxlsb.open_workbook(file_path) as wb:
                    with wb.get_sheet(sheet_name) as sheet:
                        # Read all rows from the sheet
                        data = [[cell.v for cell in row] for row in sheet.rows()]
                full_df = pd.DataFrame(data)
                
                # Use the detected header row (correct_header_row) to set column names
                full_df.columns = full_df.iloc[correct_header_row]
                # Drop the rows up to (and including) the header row and reset the index
                df = full_df.iloc[correct_header_row+1:].reset_index(drop=True)

            # Standardize column names to lowercase and strip any extra spaces
            df.columns = [str(col).strip().lower() for col in df.columns]

            print("Sheet_name:", sheet_name, "columns:", df.columns)

            if "total premium".lower() in df.columns:
                print("-----Calculating Total Premium-----")
                df["total premium"] = pd.to_numeric(df["total premium"], errors="coerce")
                return df["total premium"].sum()
        except Exception as e:
            print(f"Error processing {file_path} - {sheet_name}: {e}")

        return 0  # Default to 0 if extraction fails
    
    def process_folders(self):
        """
        Iterates through yearly folders, extracts total premium, and stores in Given_Premium.xlsx.
        """
        premium_data = []
        
        for year_folder in sorted(os.listdir(self.base_folder)):
            year_path = os.path.join(self.base_folder, year_folder)
            if not os.path.isdir(year_path):
                continue

            for insurer_file in glob.glob(os.path.join(year_path, "*")):
                if os.path.basename(insurer_file).startswith("~$"):
                    continue  # Ignore temp files

                insurer_name, ext = os.path.splitext(os.path.basename(insurer_file))
                if ext.lower() not in [".xlsx", ".xls", ".xlsb"]:  # Added .xls
                    continue  # Skip non-excel files

                print(f"Processing {insurer_name} for year {year_folder}")

                if ext.lower() == ".xlsx":
                    xl = pd.ExcelFile(insurer_file)
                    for sheet_name in xl.sheet_names if ext.lower() in [".xlsx", ".xls",".xlsb"] else xl.sheets:
                        if sheet_name.lower().startswith(("base", "reward","Base","Reward")):  # Case insensitive check
                            total_premium = self.extract_total_premium(insurer_file, sheet_name, ext.lower())
                            premium_data.append((year_folder, insurer_name, sheet_name, total_premium))

                elif ext.lower() == ".xls":
                    xl = pd.ExcelFile(insurer_file, engine="xlrd")  # Added .xls support
                    for sheet_name in xl.sheet_names if ext.lower() in [".xlsx", ".xls",".xlsb"] else xl.sheets:
                        if sheet_name.lower().startswith(("base", "reward","Base","Reward")):  # Case insensitive check
                            total_premium = self.extract_total_premium(insurer_file, sheet_name, ext.lower())
                            premium_data.append((year_folder, insurer_name, sheet_name, total_premium))

                elif ext.lower() == ".xlsb":
                    # Extract sheet names correctly
                    with pyxlsb.open_workbook(insurer_file) as wb:
                        sheet_names = list(wb.sheets)  # Get all sheet names
                    for sheet_name in sheet_names:
                        if sheet_name.lower().startswith(("base", "reward","Base","Reward")):  # Case insensitive check
                            total_premium = self.extract_total_premium(insurer_file, sheet_name, ext.lower())
                            premium_data.append((year_folder, insurer_name, sheet_name, total_premium))


        df_given = pd.DataFrame(premium_data, columns=["Year", "Insurer", "Type", "Premium"])
        df_given.to_excel(self.given_premium_file, index=False)
        print(f"Given_Premium.xlsx saved at {self.given_premium_file}")

    def refine_premium_data(self):
        """
        Aggregates base* and reward* into base and reward categories.
        """
        df = pd.read_excel(self.given_premium_file)

        # Convert Type column to lowercase before processing
        df["Type"] = df["Type"].str.lower()
        print(df)

        # Normalize Type column (base1, base2 → base; reward1, reward2 → reward)
        df["Type"] = df["Type"].apply(lambda x: "base" if x.startswith("base") else "reward")
        print(df)

        # Group by Year, Insurer, Type and sum Premium
        df_refined = df.groupby(["Year", "Insurer", "Type"], as_index=False).agg({"Premium": "sum"})
        print(df_refined)

        df_refined.to_excel(self.refined_premium_file, index=False)
        print(f"Refined_Given_Premium.xlsx saved at {self.refined_premium_file}")

    def compare_premiums(self):
        """
        Compares Refined_Given_Premium.xlsx with S3_premium.xlsx.
        """
        df_s3 = pd.read_excel(self.s3_premium_file)
        df_given = pd.read_excel(self.refined_premium_file)

        df_s3["Type"] = df_s3["Type"].str.lower()
        df_given["Type"] = df_given["Type"].str.lower()

        df_s3["Year"] = df_s3["Year"].astype(str)
        df_given["Year"] = df_given["Year"].astype(str)


        df_s3 = df_s3.groupby(["Year", "Insurer", "Type"], as_index=False).agg({"Premium": "sum"})

        # Merge on Year, Insurer, Type
        df_comparison = df_s3.merge(df_given, on=["Year", "Insurer", "Type"], how="outer", suffixes=("_S3", "_Given"))
        df_comparison["Difference"] = df_comparison["Premium_S3"].fillna(0) - df_comparison["Premium_Given"].fillna(0)

        df_comparison.to_excel(self.comparison_file, index=False)
        print(f"Comparison.xlsx saved at {self.comparison_file}")

    def run_comparison(self):
        """
        Runs the full premium comparison pipeline.
        """
        print("\nStep 1: Processing yearly folders...")
        self.process_folders()

        print("\nStep 2: Refining premium data...")
        self.refine_premium_data()

        print("\nStep 3: Comparing with S3 premium...")
        self.compare_premiums()

        print("\n✅ Comparison completed! Check the output folder for results.")

# Define Paths (Update paths based on your local setup)
base_folder = "/Users/sukrutasakoji/Downloads/Given"
s3_premium_file = "/Users/sukrutasakoji/Downloads/S3_premium_2022-23.xlsx"
output_folder = "/Users/sukrutasakoji/Downloads"

# Create an instance of PremiumComparator
comparator = PremiumComparator(base_folder, s3_premium_file, output_folder)

# Run the comparison process
comparator.run_comparison()