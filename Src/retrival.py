
import json
import asyncio
from openai import AsyncOpenAI
from DB import inserting_descriptions
from dotenv import load_dotenv

load_dotenv()
client = AsyncOpenAI()

async def main():
    with open("batch_id.txt", "r") as f:
        batch_id = f.read().strip()

    # Check status
    batch = await client.batches.retrieve(batch_id)
    print(f"Status: {batch.status}")
    print(f"Errors: {batch.errors}")
    print(f"Completed: {batch.request_counts.completed}")
    print(f"Failed: {batch.request_counts.failed}")

    if batch.status != "completed":
        print("Not ready yet, try again later!")
        return

    # Download results
    result_file = await client.files.content(batch.output_file_id)
    with open("batch_output.jsonl", "w") as f:
        f.write(result_file.text)
    print("Results downloaded!")

    # Insert to DB
    with open("batch_output.jsonl", "r") as f:
        for line in f:
            result = json.loads(line)
            custom_id = result["custom_id"]
            model, brand, type_ = custom_id.split("||")
            description = result["response"]["body"]["choices"][0]["message"]["content"]
            inserting_descriptions(model, brand, type_, "Not_Pending", description)
    
    print("All inserted to DB!")

asyncio.run(main())