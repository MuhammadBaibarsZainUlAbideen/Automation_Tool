import requests

def fetch_product_details(brand, sku, shopname, api_token=None):
    url = "https://live.icecat.biz/api"

    params = {
        "lang": "EN",
        "shopname": shopname,
        "Brand": brand,
        "ProductCode": sku,
        "content": ""   # empty = full product content
    }

    headers = {}
    if api_token:
        headers["api_token"] = api_token

    response = requests.get(url, params=params, headers=headers)

    print("Status:", response.status_code)

    if response.status_code == 200:
        return response.json()
    else:
        return response.text
print(fetch_product_details("MSI COMPUTER", "B550TMHWK", "openIcecat-live"))

