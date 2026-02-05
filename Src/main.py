import pandas as pd
#Mkaing random chnage for commit
from dotenv import load_dotenv
import os
import openai
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")
reading = pd.read_excel("Files/Product Description.xlsx")
header_list = reading.columns.tolist()
stroing = reading[header_list[2]]
empty_list = []
for i in range(0 ,len(reading)):
    response1 = openai.chat.completions.create(

        model="gpt-4o-mini",
        messages=[ {"role":"system" , "content":"Acer 146.AD406.013 is an extended service and support package designed for eligible Acer hardware. "
        "It provides enhanced post sale coverage including technical support services repair handling and service assistance for a defined coverage"
        " period. This service helps reduce downtime improves device lifecycle management and ensures reliable operational continuity. The service "
        "is non physical and does not include hardware components. Dimensions and weight are not applicable. Colour is not applicable as this is a digital service offering."
        "This Should be Your Format , Length should be exact and mimic this pattern"},
            {"role":"user","content":stroing[i]}
            ],
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

   
