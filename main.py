import httpx
import time
import random
import uuid
import pandas as pd

cookies = {
    'mg_udi': '1555498d-7fcf-4a7f-8ea7-f289f5192e42',
    'nmg_dt': 'DELIVERY_TYPE_PICKUP',
    'shopCode': '%22992301%22',
    'x_shop_type': 'ME',
    '_ym_uid': '1794245746784422749',
    '_ym_d': '1794245746',
    '_ym_isad': '2',
    'oxxfgh': 'L!5d0c73c7-4bc7-f0bb-660a-710b61e7669a%230%231800000%235000%231800000%2312840',
    'uwyii': 'de0aea8d-e40d-670b-c8f8-b113fe02ea61',
    'nmg_sp': 'Y',
    '__zzatw-magnit-rux': 'MDA0dC0cTApcfEJcdGswPi17CT4VHThHKHIzd2UqQGoiY1BhU0BcCzRbFhUwcllMD0AURD4rdTFEIyJgUFtUR1dXayELUTQ1ZhBKT01HMzg/aH0eVBw5VREPFhI2FyMSfnQnUQsQZD9IcnowN1dhMA8WEU1HFT1WUk9DKGsbcVhXL3AkF0hSfjsWa25HZ0dXTBdfQjs4WEERdU05MzUwa34mYCs5YxEOIRRDVlBGaWcVG0BIGC9LbnoyP2chZFBbJkhcVHUXXVo/KBVZcUsfbnozP14LMQ80HA==YaWRrg==',
    'spid': '1794246008557_b6e7ba636eb1c87cadf01c4e4e6dcf88_mfxu9a8m27tptam2',
    'cfidsw-magnit-rux': 'AwI6mFyXYRmA+vmgKoT1aMl3ofLKrYCeLxLvUQLb3rw9nJpsNSPW3+kHGqbT6M7zqsOx2cgF8TnUuRx/39MpYYbHwtH3+LzefHqYPhtKnKEhPrajO+VFSmHZWJq74YQz3ejvR5CHXie+Jcs9MDyEOLTBLemEj6ZmgXHA/Q==',
    'cfidsw-magnit-rux': 'AwI6mFyXYRmA+vmgKoT1aMl3ofLKrYCeLxLvUQLb3rw9nJpsNSPW3+kHGqbT6M7zqsOx2cgF8TnUuRx/39MpYYbHwtH3+LzefHqYPhtKnKEhPrajO+VFSmHZWJq74YQz3ejvR5CHXie+Jcs9MDyEOLTBLemEj6ZmgXHA/Q==',
    'gsscw-magnit-rux': 'mYwrDs8siMSxL5o/x+hqYHhT5vWw7IBfEOyvrfZXz4akFh1qC1ejpmZzIWQikP5SsKS1d1KrSHhppMfwxDVFQTqDuKb/Nu3qUKy2Bk+M2Z4ZhLA11dn4s8QIf5mhSW9WYZhhDIfneCHTy04oakI4pWvVsJheifEVQqzeB+FYD+fySYfcD+2jLgQ2SV7tCuppnpvrZmCzAigwSoHQH+YeCEJBgV0Xk8iqO2dggWvAB4birR6EomGut2P2x8Q2tbwWCGyekogAnHnWR0RtmC5zmy3h1F1A',
    'uwyiert': 'b6aa85ca-bf92-bdd2-8a59-2d2b1ee87a29',
    'fgsscw-magnit-rux': 'RcQj93645fc2ac91f29459732d3ae768560a4e3b',
}

headers = {
    'accept': 'application/json',
    'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'content-type': 'application/json',
    'origin': 'https://magnit.ru',
    'referer': 'https://magnit.ru/',
    'sec-ch-ua': '"Opera";v="139", "Chromium";v="155", "Not(A:Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36 OPR/139.0.0.0 (Edition Yx 08)',
    'x-app-version': '2026.11.9-10.3',  # Оставляем, API часто требует версию приложения
    'x-client-name': 'magnit',
    'x-device-id': str(uuid.uuid4()),   # Генерируем новый ID при каждом запуске скрипта!
    'x-device-platform': 'Web',
    'x-magnit-bff': '1',
}

json_data = {
    'categories': [
        41179,
    ],
    'includeAdultGoods': True,
    'pagination': {
        'limit': 33,
        'offset': 0,
    },
    'sort': {
        'order': 'desc',
        'type': 'popularity',
    },
    'storeCode': '992301',
    'storeType': 'dostavka',
    'catalogType': '3',
}

link = 'https://magnit.ru/webgate/v2/goods/search'

excel_filename = "data.xlsx"



def sleep():
    time.sleep(random.uniform(3,5))

def get_json(client, link, headers, json_data, cookies_data, retries=3):
    sleep()
    for i in range(retries):
        try:
            print(f"Attempt {i+1}/{retries} for {link}")
            response = client.post(link, headers=headers, json=json_data, cookies=cookies_data)

            if response.status_code in (403, 429, 500, 502, 503, 504):
                print(f"Code of answer: {response.status_code}. Retrying in 10 seconds")
                time.sleep(10)
                continue

            if response.status_code == 404:
                print(f"Page not found (404): {link}")
                return None

            response.raise_for_status()
            print("Success!")
            return response

        except httpx.TimeoutException:
            print(f"Timeout: {link}. Retrying in 2 seconds")
            time.sleep(2)
            continue

        except httpx.RequestError as e:
            # Ловим проблемы с сетью (например, пропал интернет)
            print(f"Network Error: {e}. Retrying in 5 seconds")
            time.sleep(5)
            continue

    print(f"Failed get {link} for {retries} attempts")
    return None



timeout_config = httpx.Timeout(connect=5.0, read=10.0, write=5.0, pool=5.0)
client = httpx.Client(timeout=timeout_config, http2=True)

print("=====Starting to parse data...=====\n")
res = pd.DataFrame()
for page in range(29):
    print(f"===Parsing page {page + 1}===")

    json_data['pagination']['offset'] = page*33
    items = get_json(client, link, headers, json_data, cookies)
    if items is None:
        print("Trying next page")
        continue
    else:
        items = items.json()['items']

    data_to_save = [
        {
            "ID" : item.get('id'),
            "Название" : item.get('name'),
            "Цена (руб)" : item.get('price') / 100,
            "Рейтинг" : (item.get('ratings') or {}).get("rating")
        }
        for item in items
    ]

    df = pd.DataFrame(data_to_save)
    res = pd.concat([res, df], ignore_index=True)

res.to_excel(excel_filename, index=False, engine='openpyxl')
print(f"\n=====Collected {len(res)} items=====")