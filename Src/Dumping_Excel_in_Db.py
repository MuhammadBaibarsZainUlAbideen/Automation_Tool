from DB import Table , inserting1 , deleteing, fetching_data, Dumping,inserting2
import pandas as pd
Dumping()
reading = pd.read_excel("Files/Final with description.xlsx")
batch_data = []
batch_size = 50000
count = 0
header_list = reading.columns.tolist()
for i in range(len(reading)):
    count += 1
    part_number = reading[header_list[0]][i]
    brand_name = reading[header_list[1]][i]
    descrption = reading[header_list[2]][i]
    status = "Pending"
    batch_data.append((str(part_number), str(brand_name) , str(descrption) , str(status)))
    print(count)
    if len(batch_data) == batch_size:
        inserting1(batch_data)
        batch_data.clear()

