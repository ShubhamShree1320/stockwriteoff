import pandas as pd

# Read both sheets into dataframes
sheet1_df = pd.read_excel("C:/Users/Lenovo/Desktop/RPA Write Off Input File.xlsx", sheet_name="Sheet1")
sheet2_df = pd.read_excel("C:/Users/Lenovo/Desktop/RPA Write Off Input File.xlsx", sheet_name="Sheet2")

# Merge the two dataframes based on TAG ID column
merged_df = pd.merge(sheet1_df, sheet2_df, on="TAG ID", how="left")

# Update the columns in sheet1_df with corresponding values from sheet2_df
sheet1_df["ID"] = merged_df["ID_y"]
sheet1_df["WRITE OFF TIME"] = merged_df["WRITE OFF TIME_y"]
sheet1_df["WRITE OFF DATE"] = merged_df["WRITE OFF DATE_y"]
sheet1_df["WASTAGE NUMBER"] = merged_df["WASTAGE NUMBER_y"]
sheet1_df["EXCEPTION REASON"]= merged_df["EXCEPTION REASON_y"]

sheet1_df.to_excel("updated_sheet1.xlsx", index=False)
# Optionally, you can drop the columns from sheet1_df that were added from sheet2_df
# sheet1_df.drop(["ID_y", "WRITE OFF TIME_y", "WRITE OFF DATE_y", "WASTAGE NUMBER_y"], axis=1, inplace=True)

# Optionally, fill NaN values with appropriate values if there are any
# sheet1_df.fillna(value={'ID': '', 'WRITE OFF TIME': '', 'WRITE OFF DATE': '', 'WASTAGE NUMBER': ''}, inplace=True)

# Now sheet1_df has been updated with the values from sheet2_df based on matching TAG ID
