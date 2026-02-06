import pandas as pd
from dotenv import load_dotenv
import os
import openai
import shutil
from openpyxl import load_workbook
load_dotenv()

"""User Area Only modify things in this area"""

Prompt_you_want_to_give = " manufacturer name, SKU,Technical Specs, Functions, Dimensions, " \
"                           Features and Colour, max 1000 Characters, without bullets points in " \
"                           paragraph, No Special characters, No Latin characters, Not Include The, ASCII text only"


Excel_file_whole_adress = "Files/Product Description.xlsx"

New_Excel_name = "Files/Product Description Copy.xlsx"



"""Loading Stuff Before"""
shutil.copy(Excel_file_whole_adress, f"{New_Excel_name}")


reading = pd.read_excel(Excel_file_whole_adress)
reading_copy = pd.read_excel(New_Excel_name)

header_list = reading.columns.tolist()
empty_list = []
count = 0


"""Putting Loaded Stuff in a Hashmap for better usage"""
All_Global =                   {
                "api_key": os.getenv("OPENAI_API_KEY"),
                "Brand" : reading[header_list[0]],
                "Model" : reading[header_list[1]],
                "Prompt_Header":Prompt_you_want_to_give,
                "Output_column":"Result from GPT",
                #"Output_column_Copt_thing":header_list[2],
                "count": 0
                                }



class Api_Request:
    def __init__(self, Brand , Model):
        self.Brand_Name = Brand
        self.Model = Model
        self.FinaL_string = f"{self.Brand_Name} {self.Model} {Prompt_you_want_to_give}"


    def Sending_Api_Request(self):
        

        response1 = openai.chat.completions.create(

            model="gpt-4o-mini",
            messages=[ {"role":"system" , "content":"Acer 146.AD406.013 is an extended service and support package designed for eligible Acer hardware. "
                        "It provides enhanced post sale coverage including technical support services repair handling and service assistance for a defined coverage"
                        " period. This service helps reduce downtime improves device lifecycle management and ensures reliable operational continuity. The service "
                        "is non physical and does not include hardware components. Dimensions and weight are not applicable. Colour is not applicable as this is a digital service offering."
                        "This Should be Your Format , Length should be exact and mimic this pattern"},
                        {"role":"user","content":self.FinaL_string},
            ],
        max_tokens=300,
        temperature=0.7
        )
        reply = response1.choices[0].message.content
        print(reply)
        return reply





class Main_Class_That_Will_Do_Other_things:

    def main(self):
        if All_Global["Output_column"] not in reading_copy.columns:
            reading_copy[All_Global["Output_column"]] = ""


        for i in range(0 ,len(reading)):
            #if (reading_copy.at[i,All_Global["Output_column"]]) ==  "" or (reading_copy.at[i,All_Global["Output_column_Copt_thing"]]) ==  "":  
                
                # if reading_copy.at[i,All_Global["Output_column"]] != "":
                #      continue
                All_Global["count"] += 1


                Api_request_object = Api_Request(All_Global["Brand"][i], All_Global["Model"][i])
                reply = Api_request_object.Sending_Api_Request()
                
                reading_copy.at[i,All_Global["Output_column"]] = reply
                
                if All_Global["count"] == 2:
                    with pd.ExcelWriter(New_Excel_name, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                        reading_copy[All_Global["Output_column"]].to_excel(writer,index=False,startcol=2)
                        
                        All_Global["count"]=0





Main_method = Main_Class_That_Will_Do_Other_things()
Main_method.main()
