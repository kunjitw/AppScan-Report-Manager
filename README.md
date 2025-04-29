## 基本介紹
讓你再也不用打開 AppScan 
可以直接透過匯入 xml 檔審閱弱點，並且記錄進度、筆記和截圖。

## 功能介紹
1. 透過瀏覽器介面快速瀏覽大量 AppScan XML 報告。\
   <img width="993" alt="image" src="https://github.com/user-attachments/assets/06c555c2-c044-4be4-bf27-dc860e64c5ec" />
2. 提供了篩選掃描檔的功能，快速排除低風險報告。（例如僅顯示有重大風險的掃描檔）\
   <img width="991" alt="image" src="https://github.com/user-attachments/assets/c4083f66-2c64-48e2-8f20-599fbac8b2ff" />
3. 當弱點名稱、URL、實體全部相同時（這種情況通常是版本問題，相同弱點有多個CVE）會自動折疊弱點，並只顯示嚴重程度最高的，匯出弱點時也只會匯出一個弱點。\
    <img width="601" alt="image" src="https://github.com/user-attachments/assets/63f914fe-55b1-43f5-80c4-9f10fbb72b27" />
4. 可在每個弱點（或一組相同弱點）紀錄審查狀態：「未審查」、「人工審查中」、「誤判」、「已確認弱點」。並提供筆記功能紀錄細節與其他狀況。\
   也直接從上傳或貼上剪貼簿圖片作為弱點證據，檔案將會根據掃描檔編號和弱點名稱自動命名。\
   （注意：新版本的瀏覽器可能會封鎖非 https 的剪貼簿功能，所以無法使用貼上，要手動解鎖瀏覽器限制）\
   <img width="963" alt="image" src="https://github.com/user-attachments/assets/44a5f6e6-d33a-4c7d-b443-9008db9b6c79" />
 
## 環境建置
1. 使用 python 3
2. `git clone "https://github.com/kunjitw/AppScan-Report-Manager.git"`
3. `cd AppScan-Report-Manager`
4. `pip install -r requirements.txt`
5. 建立專案資料夾 `/reports/你的專案名稱`\
   <img width="598" alt="image" src="https://github.com/user-attachments/assets/56bd1bd3-5dfc-490f-ae5c-f7de7ffb1574" />
6. 建立專案標的清單 `/reports/你的專案名稱/target.xlsx`\
   <img width="714" alt="image" src="https://github.com/user-attachments/assets/f4b1dcd6-36c3-4598-bc73-ea69c02f84e8" />
7. 編輯 target.xlsx 清單內容，必須包含三列（編號、 URL、標的名稱）\
   <img width="519" alt="image" src="https://github.com/user-attachments/assets/851a0c42-d9d6-4851-9b08-c6a4d2f5554c" />
8. 放入 .xml 掃描結果 `/reports/你的專案名稱/編號-自訂名稱.xml`\
   <img width="686" alt="image" src="https://github.com/user-attachments/assets/a2ea3f87-1ce3-4128-85d6-bd56aa5adca3" />
9. 放入 .scan 掃描檔 `/reports/你的專案名稱/編號-自訂名稱.scan`\
   （非必要，不放掃描檔除了無法直接在網頁直接開啟原始檔外，不影響其他功能）
   <img width="686" alt="image" src="https://github.com/user-attachments/assets/fddbc1c2-98ae-41fd-b976-0c31ce202d0d" />
10. 開啟網站伺服器 `python app.py`，然後選擇要開啟的 port\
    <img width="359" alt="image" src="https://github.com/user-attachments/assets/7b85d866-d475-40e8-972f-611e7f46ae25" />
11. 使用瀏覽器開啟 Web 應用 `http://127.0.0.1:port`\
    <img width="658" alt="image" src="https://github.com/user-attachments/assets/310ae8f7-9eae-4eb2-908d-cf72fe993e86" />


## 使用示範
1. 設定要排除的弱點\
   <img width="684" alt="image" src="https://github.com/user-attachments/assets/24d14ac5-73b2-413e-912f-524a9a847034" />\
   <img width="775" alt="image" src="https://github.com/user-attachments/assets/8de78cb9-28fe-4669-8373-d409746f8ea5" />
3. 選擇要處理的專案\
   <img width="684" alt="image" src="https://github.com/user-attachments/assets/0936a205-4012-40d7-942c-fa3ff095275f" />
4. 首先我想確認掃描狀況，先選擇所有報告（含異常）\
   <img width="1024" alt="image" src="https://github.com/user-attachments/assets/f2ea0516-b6a2-4b48-a260-0f5e3b94abf4" />\
   打叉為有掃描檔，但是掃描失敗｜禁止符號為找不到掃描檔｜打勾為正常\
   <img width="997" alt="image" src="https://github.com/user-attachments/assets/40b7f298-0f78-48fa-ad0b-f92bcc0e655b" />
6. 接下來我只想看有中風險以上的弱點，這時選中風險及以上\
   <img width="1017" alt="image" src="https://github.com/user-attachments/assets/335a46a5-73c3-4571-8cfd-ffb369fe4bfb" />\
   這時選都會幫你把有中風險弱點以上的報告挑出來置頂\
   <img width="996" alt="image" src="https://github.com/user-attachments/assets/bb6424a8-86b0-4b04-858c-406ce340b1d4" />
7. 點選網址觀看掃描檔詳細資料\
   <img width="996" alt="image" src="https://github.com/user-attachments/assets/c77798fd-4021-4154-95d2-1eb8c404399c" />
8. 可以看到掃描檔的一些資本資訊\
   （其中 標的名稱 和 目標URL 是由你專案內手動新增的 target.xlsx 抓取，其他則為 .xml 掃描檔內的原始資料）
   <img width="921" alt="image" src="https://github.com/user-attachments/assets/56d9a5dd-459b-4538-ab50-721d30862dfc" />
9. 往下拉可以看到所有弱點，並且可以使用過濾器篩選掉你不想看的內容\
   <img width="909" alt="image" src="https://github.com/user-attachments/assets/47c298ae-7832-4d10-b2c5-f5be324b8d3a" />
10. 接著開始驗證弱點，並且記錄驗證資訊\
    狀態：（只有設定為 已確認弱點 的，匯出弱點時才會被匯出）\
    已完成全部截圖：（不會影響匯出弱點的功能，主要是讓你紀錄這個弱點的截圖你是否都全部完成了）\
    上傳與貼上按鈕：可以上傳或是貼上剪貼簿的圖片，並且自動幫你命名成好辨識的格式，檔案會保存在\
    筆記：（不會影響匯出弱點的功能，單純給你紀錄用的，輸出弱點時可以選擇是否要輸出該弱點的筆記內容）\
    `\AppScan-Report-Manager\data\你的專案名稱\screenshots`
    <img width="911" alt="image" src="https://github.com/user-attachments/assets/3ad84fe3-92f2-42e9-89d6-56f8d015a5c8" />
12. 當你把一個掃描檔判讀完成後，你可以回到選擇報告清單，把判讀完成的欄位打勾\
    （這不會影響匯出或是其他功能，單純幫助你記憶而已）\
    <img width="929" alt="image" src="https://github.com/user-attachments/assets/61ac76d5-674b-4f5d-99a0-a7bbd64bbd46" />
13. 最後你可以隨時使用 匯出確認弱點 的功能，來匯出所有 狀態為已確認弱點 的弱點\
   （基本上都使用 Excel 匯出，JSON 未來計劃匯出的弱點可直接接入其他系統使用）\
    <img width="1005" alt="image" src="https://github.com/user-attachments/assets/34ca93c1-4ba4-4616-a02c-b071fe6f39c1" />
14. 匯出已確認弱點 Excel 格式\
    <img width="1771" alt="image" src="https://github.com/user-attachments/assets/328815c4-8dec-4ae2-8e23-84ffb9fb4ae9" />


## 其他示範
1. 如果你想刪除你的判讀紀錄，可前往 `AppScan-Report-Manager\data\你的專案名稱` 把 `vulnerability_status.json` 刪除

## 目前問題:
  - 專案名稱是由專案資料夾決定的無法更改
  - 資料夾改名後，相關資料不會自動轉移，改名後會變成一份全新沒有紀錄的專案
  - 深色模式只有專案頁面才有
  - 深色模式下某些字體配色失效
  - 驗證按鈕無法同時執行，一次只能處理一個弱點
  - 以單人使用為前提設計，沒有對多人同時使用進行相關的測試以及最佳化
  - 目前 XML 使用 AppScan 10.7.0 版本匯出，無測試其他本版 XML 可用性
  - 弱點不會自動擴充，目前一個專案只能輸出 7 個弱點
   
## 未來規劃:
  - 加入英文版介面
  - 加入 Acuntix 掃描檔讀取功能
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
