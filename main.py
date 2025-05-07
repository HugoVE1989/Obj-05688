import openpyxl
import pandas as pd
import xlwings as xw
import os as os
from functions import quantity_check_layer, data_present


folder_path = r"data_folder"
files_in_folder = os.listdir(folder_path)

excel_file_paths = []
for excel_file in files_in_folder:
    file_path = os.path.join(folder_path, excel_file)
    excel_file_paths.append(file_path)

df_folder = pd.DataFrame()
df_dimension_group = pd.DataFrame()

df_total = pd.DataFrame()
for excel_file_path in excel_file_paths:
    df = pd.read_csv(excel_file_path, sep=";")

    df_total = pd.concat([df_total, df], ignore_index=True)

#Split the data_folder in the dataframe
df_split = pd.DataFrame()
df_split[["Model", "Object", "OTL", "Laagdikte"]] = df_total['Folder'].str.split('\\', expand=True)
df_total = pd.concat([df_total, df_split], axis=1)

#Only use the following
df_total = df_total.reindex(columns=["Model", "Object", "OTL", "Laagdikte",
                                     "Custom 1 Name", "Custom 1", "Custom 1 UOM",
                                     "Custom 2 Name", "Custom 2", "Custom 2 UOM",
                                     "Custom 4 Name", "Custom 4", "Custom 4 UOM"])

df_results = df_total.rename(columns={  "Custom 1 Name": "ROCO_CVO_Volume_Total", "Custom 1": "Volume", "Custom 1 UOM": "m3",
                                        "Custom 2 Name": "ROCO_CAR_Sectional_Area", "Custom 2": "Area", "Custom 2 UOM": "m2",
                                        "Custom 4 Name": "ROCO_CLE_Length_Total", "Custom 4": "Length", "Custom 4 UOM" : "m"
                                        })

xw.view(df_results)
df_selection = quantity_check_layer(df=df_results, otl="OTL-1003-Bitumineuze Toplaag")