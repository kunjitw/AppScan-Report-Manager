## 基本介紹
讓你再也不用打開 AppScan 就可以審閱報告弱點，並且記錄進度、筆記和截圖。

## 功能介紹
1. 透過瀏覽器介面快速瀏覽大量 AppScan XML 報告。
2. 提供了篩選掃描檔的功能，快速排除低風險報告。
3. 可在每個弱點（或一組相同弱點）紀錄審查狀態：「未審查」、「人工審查中」、「誤判」、「已確認弱點」。並提供筆記功能紀錄細節與其他狀況。
4. 直接從介面 **上傳** 圖片檔案作為弱點證據，也支援從剪貼簿**貼上**圖片快速新增截圖。
   （注意：新版本的瀏覽器可能會封鎖非 https 的剪貼簿功能，所以無法使用貼上，要手動解鎖瀏覽器限制）
5. 當弱點名稱、URL、實體全部相同時（這種情況通常是版本問題，相同弱點有多個CVE）會自動折疊弱點，並只顯示嚴重程度最高的，匯出弱點時也只會匯出一個弱點。
 
## 環境建置
1. 使用 python 3
2. `git clone "https://github.com/kunjitw/AppScan-Report-Manager.git"`
3. `cd AppScan-Report-Manager`
4. `pip install -r requirements.txt`
5. 建立專案資料夾 `/report/你的專案名稱`
6. 建立專案標的清單 `/report/你的專案名稱/target.xlsx`
7. 放入 .xml 掃描結果 `/report/你的專案名稱/編號-自訂名稱.xml`
8. 放入 .scan 掃描檔 `/report/你的專案名稱/編號-自訂名稱.scan`
9. 開啟網站伺服器 `python app.py`，然後選擇要開啟的 port
    <img width="359" alt="image" src="https://github.com/user-attachments/assets/7b85d866-d475-40e8-972f-611e7f46ae25" />
10. 使用瀏覽器開啟 Web 應用 `http://127.0.0.1:port`
    <img width="658" alt="image" src="https://github.com/user-attachments/assets/310ae8f7-9eae-4eb2-908d-cf72fe993e86" />


## 使用示範
1. 設定要排除的弱點
   <img width="684" alt="image" src="https://github.com/user-attachments/assets/24d14ac5-73b2-413e-912f-524a9a847034" />
   <img width="775" alt="image" src="https://github.com/user-attachments/assets/8de78cb9-28fe-4669-8373-d409746f8ea5" />
2. 選擇要處理的專案
   <img width="684" alt="image" src="https://github.com/user-attachments/assets/0936a205-4012-40d7-942c-fa3ff095275f" />
4. 首先我想確認掃描狀況
6. 我只想看有中風險以上的弱點，這時選中風險以上
7. 

## 目前問題:
  - 專案名稱是由專案資料夾決定的無法更改
  - 資料夾改名後，相關資料不會自動轉移，改名後會變成一份全新沒有紀錄的專案
  - 深色模式只有專案頁面才有
  - 深色模式下某些字體配色失效
  - 驗證按鈕無法同時執行，一次只能處理一個弱點
  - 以單人使用為前提設計，沒有對多人同時使用進行相關的測試以及最佳化
  - 目前 XML 使用 AppScan 10.7.0 版本匯出，無測試其他本版 XML 可用性
   
## 未來規劃:
  - 加入英文版介面加入 Acuntix 掃描檔讀取功能
  - 在專案頁面加入上下頁按鈕，加入勾選按鈕（是否已檢測完畢）
  - 打開清單後 自動跳到目前瀏覽專案，並且高亮顯示

## 修改紀錄
- 2025/04/28 Rafe
  - 統一了上傳和修改圖片的提示視窗
- 2025/04/27 Rafe
  - 加入了 json 匯出（包含圖片 base64）
  - 加入了圖片備註功能
  - 加入了可選擇常用圖片備註功能
  - 加入了檔案 `json解析範例.html` 可讀取匯出的 json 檔案，並呈現內容
