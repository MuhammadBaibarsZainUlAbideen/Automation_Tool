import pandas as pd
from dotenv import load_dotenv
import os
import openai
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")
reading = pd.read_excel("Files/Product Description.xlsx")
stroing = reading["Prompt 1"]
empty_list = []
for i in range(0 ,len(reading)):
    response1 = openai.chat.completions.create(

        model="gpt-4o-mini",
        messages=[{"role":"user","content":stroing[i]}],
        max_tokens=50,
        temperature=0.7
    )
    reply = response1.choices[0].message.content
    print(reply)
    reading.at[i,"Product Description (Result from Chat GPT)"] = reply
    empty_list.append(reply)


reading.to_excel("Ouput/o.xlsx",index=False)
with pd.ExcelWriter("Files/Product Description.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    reading["Product Description (Result from Chat GPT)"].to_excel(writer,index=False,startcol=3)