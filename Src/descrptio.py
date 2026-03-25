from DB import fetching_data2
import pandas as pd
data = fetching_data2()
df = pd.DataFrame(data, columns=['model', 'brand', 'type', 'status', 'description'])
df.to_excel("Set 3.0.xlsx", index=False)