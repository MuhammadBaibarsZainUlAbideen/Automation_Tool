import pandas as pd
from dotenv import load_dotenv
import os
import openai
import shutil
import asyncio
load_dotenv()

# ---------------- User Area ----------------
Prompt_you_want_to_give = (
    "manufacturer name, SKU,Technical Specs, Functions, Dimensions, "
    "Features and Colour, max 1000 Characters, without bullets points in "
    "paragraph, No Special characters, No Latin characters, Not Include The, ASCII text only"
)

Excel_file_whole_adress = "Files/Product Description.xlsx"
New_Excel_name = "Files/Product Description Copy.xlsx"
shutil.copy(Excel_file_whole_adress, f"{New_Excel_name}")
reading = pd.read_excel(Excel_file_whole_adress)
reading_copy = pd.read_excel(New_Excel_name)

header_list = reading.columns.tolist()

All_Global = {
    "api_key": os.getenv("OPENAI_API_KEY"),
    "Brand": reading[header_list[0]],
    "Model": reading[header_list[1]],
    "Prompt_Header": Prompt_you_want_to_give,
    "Output_column": "Result from GPT",
}

openai.api_key = All_Global["api_key"]

# ---------------- API Class ----------------
class Api_Request:
    def __init__(self, Brand, Model):
        self.Brand_Name = Brand
        self.Model = Model
        self.FinaL_string = f"{self.Brand_Name} {self.Model} {Prompt_you_want_to_give}"

    async def send_request(self, semaphore):
        async with semaphore:
            return await asyncio.to_thread(self._sync_request)

    def _sync_request(self):
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Acer 146.AD406.013 is an extended service and support package designed for eligible Acer hardware. "
                        "It provides enhanced post sale coverage including technical support services repair handling and service assistance for a defined coverage "
                        "period. This service helps reduce downtime improves device lifecycle management and ensures reliable operational continuity. The service "
                        "is non physical and does not include hardware components. Dimensions and weight are not applicable. Colour is not applicable as this is a digital service offering. "
                        "This Should be Your Format , Length should be exact and mimic this pattern"
                    ),
                },
                {"role": "user", "content": self.FinaL_string},
            ],
            max_tokens=300,
            temperature=0.7,
        )
        reply = response.choices[0].message.content
        print(reply)
        return reply

# ---------------- Main Processing ----------------
class Main_Class:
    def __init__(self, batch_size=5, concurrency=5):
        self.semaphore = asyncio.Semaphore(concurrency)  # concurrent API calls
        self.batch_size = batch_size

    async def process_row(self, i):
        api_obj = Api_Request(All_Global["Brand"][i], All_Global["Model"][i])
        reply = await api_obj.send_request(self.semaphore)
        reading_copy.at[i, All_Global["Output_column"]] = reply

    async def main(self):
        if All_Global["Output_column"] not in reading_copy.columns:
            reading_copy[All_Global["Output_column"]] = ""

        total_rows = len(reading)
        for start in range(0, total_rows, self.batch_size):
            end = min(start + self.batch_size, total_rows)
            tasks = [self.process_row(i) for i in range(start, end)]
            await asyncio.gather(*tasks)

            # Save after each batch
            reading_copy.to_excel(New_Excel_name, index=False)
            print(f"Saved rows {start} to {end-1}.")

# ---------------- Run ----------------
Main_method = Main_Class(batch_size=5, concurrency=5)
asyncio.run(Main_method.main())
