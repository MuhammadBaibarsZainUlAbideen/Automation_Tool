import pandas as pd
from dotenv import load_dotenv
import os
from openai import AsyncOpenAI
import shutil
from openpyxl import load_workbook
import asyncio
import os
from openai import AsyncAzureOpenAI

from DB import fetching_data1,inserting_descriptions,indexation

load_dotenv()
indexation()

endpoint = "https://foundry-ai-prd-eus2-01.cognitiveservices.azure.com/"
model_name = "gpt-4o-mini"
deployment = "gpt-4o-mini"
subscription_key = os.getenv("subscription_key")
api_version = "2024-12-01-preview"
client = AsyncAzureOpenAI(
    api_version=api_version,
    azure_endpoint=endpoint,
    api_key=subscription_key,
)

"""User Area Only modify things in this area"""

Prompt_you_want_to_give = " manufacturer name, SKU,Technical Specs, Functions, Dimensions, " \
"                           Features and Colour, max 1000 Characters, without bullets points in " \
"                           paragraph, No Special characters, No Latin characters, Not Include The, ASCII text only"











# """Putting Loaded Stuff in a Hashmap for better usage"""
# All_Global =                   {
#                 "api_key": os.getenv("OPENAI_API_KEY"),
#                 "Brand" : reading[header_list[1]],
#                 "Model" : reading[header_list[0]],
#                 "Type"  : reading[header_list[2]],
#                 "Prompt_Header":Prompt_you_want_to_give,
#                 "Output_column":"Result from GPT",
#                 #"Output_column_Copt_thing":header_list[2],
#                 "count": 0
#                                 }



class Api_Request:
    semaphore = asyncio.Semaphore(30)
    def __init__(self, Brand , Model , type):
        self.Brand_Name = Brand
        self.Model = Model
        self.type = type
        self.FinaL_string = f"{self.Brand_Name} {self.Model} {self.type} {Prompt_you_want_to_give}"
        

    
    async def Sending_Api_Request(self):

        
        async with Api_Request.semaphore:
            try:
                response1 = await client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content":  f"You are a product description writer for U.S. federal GSA procurement listings. Write accurate descriptions based only on the product information provided Must mention the SKU/partnumber and expand the descrption and mention the provided decrption as well that will be provided to you, expand it as well that get all the information. Do not infer, assume, or add any information not explicitly provided. Do not mention fields that are not applicable or not provided such as dimensions or color or Techinical Specs, functions, features for software and license products. Write in paragraph format, maximum 1000 characters including spaces, using only standard ASCII characters. Do not use bullet points, special symbols, or marketing language. Maintain a formal, neutral, and professional tone. Mention any company names only as the manufacturer or source of the product, not as the provider of the service, unless explicitly stated",
                        },
                        {
                            "role": "user",
                            "content": self.FinaL_string,
                        }
                    ],
                    max_tokens=300,
                    temperature=0.7,
                    top_p=1.0,
                    model=deployment
                    )

                reply = response1.choices[0].message.content
                print(reply)
                
                return reply
            except:
                error = "Error"
                print(error)
                return error






class Main_Class_That_Will_Do_Other_things:
    async def main(self):
        data = fetching_data1()
        count = 0 
        tasks = []
        pending_rows = []  
        btach_insert = []


        for i in range(len(data)):
           
        
        
            if data[i][3] == "Pending":
                count +=1
               
                Api_request_object = Api_Request(data[i][1], data[i][0], data[i][2])
                tasks.append(Api_request_object.Sending_Api_Request())
                pending_rows.append(data[i])  # save the row at same index as task
                
            if (len(tasks) >= 100):
                print(count)
                results = await asyncio.gather(*tasks)
                for idx,result in enumerate(results):
                    row = pending_rows[idx]
                    btach_insert.append((result, "Not_Pending", row[0], row[1], row[2]))
                        
                inserting_descriptions(btach_insert)
                btach_insert.clear()
             
                tasks.clear()        
                pending_rows.clear()
                
        if tasks:
            results = await asyncio.gather(*tasks, return_exceptions=True)
            for idx, result in enumerate(results):
                if isinstance(result, Exception):
                    print(f"Failed: {result}")
                    continue
                row = pending_rows[idx]
                btach_insert.append((result, "Not_Pending", row[0], row[1], row[2]))

            inserting_descriptions(btach_insert)
            btach_insert.clear()
           
                       
                    

        # df = pd.DataFrame(batch_data, columns=['model', 'brand', 'type', 'status', 'description'])
        # df.to_excel("output_smart_prompt.xlsx", index=False)



                        # new_count = 0
                        # print("Nah")
                        # if len(batch_data) >= batch_size:
                        #     print("Hell2")

              
                        #     batch_data.clear()
                        #     break

                        







                #     new_count = 0
                
                # if All_Global["count"] == Save_your_data_to_disk_after_how_many_replies:
                #     with pd.ExcelWriter(New_Excel_name, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                #         reading_copy[All_Global["Output_column"]].to_excel(writer,index=False,startcol=2)
                        
                #         All_Global["count"]=0





Main_method = Main_Class_That_Will_Do_Other_things()
asyncio.run(Main_method.main())

