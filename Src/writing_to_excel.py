from DB import  fetching_data1
import pandas as pd
Excel_file_whole_adress = "Files/Product Description copy 8.xlsx"

db_data = fetching_data1()  # [(brand, model, description), ...]
print(db_data)
df = pd.DataFrame(db_data, columns=["Brand", "Model", "Type" , "Result from GPT"])
df.to_excel(Excel_file_whole_adress, index=False)


