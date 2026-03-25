import json
import asyncio
from openai import AsyncOpenAI
from DB import fetching_data1
from dotenv import load_dotenv
from DB import fetching_data1,inserting_descriptions
load_dotenv()
client = AsyncOpenAI()

Prompt_you_want_to_give = "manufacturer name, SKU, Technical Specs, Functions, Dimensions, Features and Colour, max 1000 Characters, without bullet points in paragraph, No Special characters, No Latin characters, Not Include The, ASCII text only"

async def main():
    data = fetching_data1()
    
    # Only take first 50,000 pending rows
    requests = []
    count = 0
    for row in data:
        if row[3] == "Pending":
            request = {
                "custom_id": f"{row[0]}||{row[1]}||{row[2]}",  # model||brand||type
                "method": "POST",
                "url": "/v1/chat/completions",
                "body": {
                    "model": "gpt-4o-mini",
                    "messages": [
                        {"role": "system", "content": "You are a product description writer. Write accurate descriptions based only on the product information provided. Do not include any information not provided. Format: manufacturer name, product name, SKU, Technical Specs, Functions, Dimensions, Features, and Color. Write in paragraph format, maximum 1000 characters including spaces, using only standard ASCII characters. Do not use bullet points or special symbols. Maintain a clear and professional tone."},
                        {"role": "user", "content": f"{row[1]} {row[0]} {row[2]} {Prompt_you_want_to_give}"}
                    ],
                    "max_tokens": 300,
                    "temperature": 0.7
                }
            }
            requests.append(request)
            count += 1
            if count == 40000:
                break
    
    # Write to JSONL file
    with open("batch_input.jsonl", "w") as f:
        for req in requests:
            f.write(json.dumps(req) + "\n")
    print(f"Created {len(requests)} requests in batch_input.jsonl")

    # Upload file to OpenAI
    with open("batch_input.jsonl", "rb") as f:
        uploaded_file = await client.files.create(file=f, purpose="batch")
    print(f"File uploaded: {uploaded_file.id}")

    # Submit batch
    batch = await client.batches.create(
        input_file_id=uploaded_file.id,
        endpoint="/v1/chat/completions",
        completion_window="24h"
    )
    print(f"Batch submitted! ID: {batch.id}")

    # Save batch ID
    with open("batch_id.txt", "w") as f:
        f.write(batch.id)
    print("Batch ID saved to batch_id.txt — run retrieve.py tomorrow!")

asyncio.run(main())