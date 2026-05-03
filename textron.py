import requests
import pandas as pd
import os
import base64


INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output_check.xlsx"
IMAGE_FOLDER = "images"

os.makedirs(IMAGE_FOLDER, exist_ok=True)

ENCODED_URL = "aHR0cHM6Ly90ZXh0cm9uc3BlY2lhbGl6ZWR2ZWhpY2xlcy5teS5zaXRlLmNvbS9UZXh0cm9uR1NFT3JlL3Mvcy9zZnNpdGVzL2F1cmE="

def get_url():
    return base64.b64decode(ENCODED_URL).decode()


def get_headers():
    return {
        'accept': '*/*',
        'accept-language': 'en-US,en;q=0.8',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'origin': 'https://textronspecializedvehicles.my.site.com',
        'priority': 'u=1, i',
        'referer': 'https://textronspecializedvehicles.my.site.com/TextronGSEStore/s/global-search/660-1-3028',
        'sec-ch-ua': '"Brave";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'sec-gpc': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36',
        'x-b3-sampled': '1',
        'x-b3-spanid': 'd25d37e83b8c45c7',
        'x-b3-traceid': '1749e8f42d863ff0',
        'x-sfdc-lds-endpoints': 'ApexActionController.execute:B2BCommerce_GetInfo.getCartSummary, ApexActionController.execute:B2BCommerce_SearchController.productSearch',
        'x-sfdc-page-cache': 'fe506b3a34078056',
        'x-sfdc-page-scope-id': 'c288bc20-9f73-47af-9437-23290eed16c7',
        'x-sfdc-request-id': '10459983000000e8b7'}


def first_api(part_number):

    url = get_url()

    payload = f"""message={{"actions":[
    {{"id":"1;a","descriptor":"aura://ApexActionController/ACTION$execute","params":{{"classname":"B2BCommerce_SearchController","method":"productSearch","params":{{"searchQuery":"{{\\"searchTerm\\":\\"{part_number}\\"}}"}}}}}}
    ]}}"""

    response = requests.post(url, headers=get_headers(), data=payload)

    try:
        data = response.json()
        products = data["actions"][0]["returnValue"]["returnValue"]["productsPage"]["products"]
    except:
        return None, None, None

    if not products:
        return None, None, None

    for product in products:
        try:
            code = product["fields"]["ProductCode"]["value"]

            if code.strip().lower() == part_number.strip().lower():
                return (
                    product.get("id"),
                    code,
                    product["defaultImage"].get("url", "")
                )
        except:
            continue

    return None, None, None


def second_api(product_id):

    url = get_url()

    payload = f"""message={{"actions":[
    {{"id":"2;a","descriptor":"aura://ApexActionController/ACTION$execute","params":{{"classname":"B2BCommerce_PriceAndInventoryController","method":"getPriceAndInventoryInfo","params":{{"productId":"{product_id}","quantity":1}}}}}}
    ]}}"""

    response = requests.post(url, headers=get_headers(), data=payload)

    try:
        item = response.json()["actions"][0]["returnValue"]["returnValue"]["items"][0]
        return item.get("dealerPrice"), item.get("inventoryQuantity")
    except:
        return None, None


def download_image(image_url, part_number):

    if not image_url or "default-product-image" in image_url:
        return "No Image"

    base = base64.b64decode("aHR0cHM6Ly90ZXh0cm9uc3BlY2lhbGl6ZWR2ZWhpY2xlcy5teS5zaXRlLmNvbQ==").decode()
    full_url = base + image_url

    try:
        content = requests.get(full_url).content
        path = os.path.join(IMAGE_FOLDER, f"{part_number}.jpg")

        with open(path, "wb") as f:
            f.write(content)

        return full_url

    except:
        return ""


def main():

    df = pd.read_excel(INPUT_FILE)
    results = []

    for part_number in df["search input"]:

        print("Processing:", part_number)

        product_id, product_code, image_url = first_api(str(part_number))

        if not product_id:
            results.append([part_number, "Not Found", "-", "-", "-"])
            continue

        if str(part_number).strip() != str(product_code).strip():
            results.append([part_number, product_code, "Mismatch", "-", "-"])
            continue

        price, inventory = second_api(product_id)

        stock = "In Stock" if inventory and inventory > 0 else "Out of Stock"
        image_link = download_image(image_url, part_number)

        results.append([
            part_number,
            product_code,
            stock,
            price,
            image_link
        ])

    output_df = pd.DataFrame(results, columns=[
        "Search Input",
        "Product Code",
        "Inventory",
        "Price",
        "Image URL"
    ])

    output_df.to_excel(OUTPUT_FILE, index=False)

    print("✅ Done")


if __name__ == "__main__":
    main()