import pandas as pd
import requests

while True:
# === 設定 ===
    ACCESS_TOKEN = 'SHOPLINE_API_KEY'
    API_BASE_URL = 'https://open.shopline.io/v1/'
    file = input('請輸入完成名單(含訂單編號/訂購人/總計(對帳金額)): ')
    df = pd.read_excel(file)

    headers = {
        'Authorization': f'Bearer {ACCESS_TOKEN}',
        'Content-Type': 'application/json'
    }

# 用來儲存每筆處理結果
    results = []

    for name, order_number, price in zip(df['訂購人'], df['訂單編號'], df['總計(對帳金額)']):
        if pd.isna(order_number) or pd.isna(price):
            continue

        order_number = str(order_number).replace('#', '').strip()

        # 👉 這裡把 price 清洗成 int
        try:
            price = int(float(str(price).replace(',', '').strip()))
        except ValueError:
            results.append({'訂購人': name, '訂單編號': order_number,
                            '金額': price, '狀態': '❌ 金額格式錯誤'})
            continue
            
        search_url = f"{API_BASE_URL}/orders/search?order_number={order_number}"
        response = requests.get(search_url, headers=headers)

        if response.status_code != 200:
            results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '❌ 查詢失敗'})
            continue

        data = response.json()
        if 'items' in data and len(data['items']) > 0:
            order_id = data['items'][0]['id']
            api_price = data['items'][0]['total']['dollars']

            if api_price != price:
                results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '❌ 金額不符'})
                continue

            # Step 1: 更新訂單狀態為 completed
            patch_status_url = f"{API_BASE_URL}/orders/{order_id}/status"
            patch_status_payload = {"status": "completed"}
            patch_status_response = requests.patch(patch_status_url, headers=headers, json=patch_status_payload)

            if patch_status_response.status_code == 200:
                # Step 2: 更新為 collected
                patch_delivery_status_url = f"{API_BASE_URL}/orders/{order_id}/order_delivery_status"
                patch_delivery_status_payload = {"status": "collected"}
                patch_delivery_status_response = requests.patch(patch_delivery_status_url, headers=headers, json=patch_delivery_status_payload)

                if patch_delivery_status_response.status_code == 200:
                    results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '✅ 已完成 + 已取貨'})
                else:
                    results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '✅ 已完成，❌ 已取貨失敗'})
            else:
                results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '❌ 更新失敗'})
        else:
            results.append({'訂購人': name, '訂單編號': order_number, '金額': price, '狀態': '❌ 沒有找到對應的訂單'})

    # === 將結果寫入 Excel 檔案 ===
    result_df = pd.DataFrame(results)
    result_df.to_excel('回傳結果.xlsx', index=False)
    print("✅ 回傳結果.xlsx")
    print('\n')
input('請按enter鍵結束')
