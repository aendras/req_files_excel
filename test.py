import pandas as pd
df = pd.read_excel("DFMC M18-3 FBL SRS Rev 2.5_F002_ET0_20241008.xlsx", sheet_name = "6_1 Default value", engine = 'openpyxl')
print(df)