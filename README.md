# AppScan-Report-Manager

一款基於 Web 的工具，旨在簡化 [HCL AppScan Standard](https://www.hcltechsw.com/appscan) 產生的 XML 漏洞報告審查流程。\
提供了審查狀態追蹤、筆記、截圖管理、自動排除規則以及匯出已驗證弱點之功能。

---

## ✨ 主要功能

*   **🌐 Web 化報告瀏覽:**
    *   透過瀏覽器介面快速瀏覽大量 AppScan XML 報告。\
        <img width="834" alt="image" src="https://github.com/user-attachments/assets/a9028d95-1505-4efb-8730-48021f649759" />
    *   提供了篩選掃描檔的功能，快速排除低風險報告。\
        <img width="1125" alt="image" src="https://github.com/user-attachments/assets/78c531ca-b872-40f2-b34e-66f0efabfc23" />
        <img width="935" alt="image" src="https://github.com/user-attachments/assets/ec589256-e566-43e9-ac51-09a64c773c28" />
    *   視覺化顯示掃描摘要和掃描統計數據，並提供快捷鍵，幫助你開啟對印的掃描原始檔（本功能會將掃描檔開啟在 Server 端）。\
        <img width="654" alt="程式主介面" src="https://github.com/user-attachments/assets/46555b2f-489a-4e1a-8ad1-efc9181162be" />
    *   每個弱點以獨立卡片展示，包含嚴重性、URL、受影響實體（參數、Cookie 等）、AppScan 判斷原因及相關 HTTP 流量。\
        <img width="614" alt="單一問題卡片" src="https://github.com/user-attachments/assets/44349654-93cb-4c45-8061-9bbed6c39445" />
    *   可在每個弱點（或一組相同弱點）紀錄審查狀態：「未審查」、「人工審查中」、「誤判」、「已確認弱點」。並提供筆記功能紀錄細節與其他狀況。\
        <img width="1088" alt="image" src="https://github.com/user-attachments/assets/cde5a591-86c8-4ced-aae3-39e6c4ff4f4a" />
    *   直接從介面**上傳**圖片檔案作為弱點證據，也支援從剪貼簿**貼上**圖片快速新增截圖。
    *   檔案會自動命名為（標的編號-弱點來源-弱點名稱-弱點URL-弱點實體-流水號），
    *   並保存在 `data/專案名稱/screenshots`
    *   刪除圖片時圖片會被移動至`_trash_screenshots`資料夾內。\
        <img width="1090" alt="image" src="https://github.com/user-attachments/assets/29fd97b0-2b03-4c27-a3dd-38bcbaf01e22" />
    *   標記某問題是否已**完成全部截圖**，已確認收集完弱點的所有證據。\
        <img width="1094" alt="image" src="https://github.com/user-attachments/assets/578e7ad6-cb73-4884-aeae-373e83953ece" />
    *   當弱點名稱、URL、實體全部相同時（這種情況通常是版本問題，有多個CVE），自動折疊弱點，並只顯示嚴重程度最高的，匯出弱點時也只會匯出一個弱點。\
        <img width="807" alt="image" src="https://github.com/user-attachments/assets/8a53b7dc-8608-4d3d-b3c0-b9950acb3e84" />
    *   可依照專案分類掃描檔。\
        <img width="900" alt="專案選擇頁面" src="https://github.com/user-attachments/assets/84bdaff9-67ff-4c9e-885f-b2a5177c99b2" />

*   **⚙️ 自動排除規則:**
    *   設定全域規則，根據「弱點類型」或結合「實體名稱」（開頭/包含）\
        <img width="976" alt="排除規則設定" src="https://github.com/user-attachments/assets/6e0f7a17-0f81-44ed-9319-b4695cdd7fff" />
    *   自動將符合條件的問題標記為「已自動排除」，減少重複性的人工誤判標記工作。\
        <img width="828" alt="自動排除無用弱點" src="https://github.com/user-attachments/assets/370ee9ae-ff31-4eea-88a4-aea50d838a93" />

*   **🔍 輔助驗證（Selenium）（本會將瀏覽器開啟在 Server 端）:**
    *   針對「外部連結」或「已知弱點元件」類型的問題，可嘗試使用 驗證按鈕\
        <img width="1072" alt="image" src="https://github.com/user-attachments/assets/cbe8dc95-838e-4f0f-9962-ff5b832d0c0a" />
    *   本功能會開啟瀏覽器，自動在目標原始碼中查找相關的外部域名、元件版本號或元件名稱\
        <img width="968" alt="image" src="https://github.com/user-attachments/assets/55676a82-62e0-450f-aa97-0776fccd21ac" />
    *   瀏覽器會自動滾動至目標，方便截圖證據。\
        <img width="1032" alt="image" src="https://github.com/user-attachments/assets/32377a41-3def-4f15-b636-43257c808f80" />

*   **📑 匯出 Excel:**
    *   匯出**已確認弱點**:
        *   可選擇是否**合併**相同類型的弱點。
        *   可選擇是否**包含**問題筆記。
        <img width="1492" alt="image" src="https://github.com/user-attachments/assets/8c3d6076-122c-435e-967a-0ee56ba6fa03" />
    *   匯出**所有筆記**: 包含所有報告中的所有問題及其筆記、狀態等詳細資訊（不管是否有筆記的會匯出，等於匯出所有弱點，包含已確認、誤判、自動排除...）。\
        <img width="1478" alt="image" src="https://github.com/user-attachments/assets/8247fb13-3bdc-4a80-ac3a-e7bd86b95c55" />

    *   匯出**異常報告**: 列出掃描失敗、解析錯誤或在目標清單中但找不到對應 XML 檔案的報告，方便重新掃描。\
        <img width="1136" alt="image" src="https://github.com/user-attachments/assets/f2d04d91-ed34-4f16-814c-267ec0baeeae" />
        <img width="1110" alt="image" src="https://github.com/user-attachments/assets/a3ad0de4-b507-45b6-badf-9a76a9af1da6" />



*   **➕ 手動新增弱點:**
    *   可在特定報告中手動添加不在 AppScan 掃描結果中的弱點，並進行狀態管理、筆記和截圖。\
        <img width="1129" alt="image" src="https://github.com/user-attachments/assets/f4c64acb-25d2-4e24-911b-76d6bef928a4" />
        <img width="728" alt="image" src="https://github.com/user-attachments/assets/0234c1ef-7b41-4312-acf8-684a21d3bfcb" />
    *   你可以在 `_internal\data\weakness_list.txt` 內加入常用弱點名稱，手動新增弱點時會變成下拉選單，可快速選取。\
        <img width="875" alt="image" src="https://github.com/user-attachments/assets/797392de-e521-4b93-8242-b6a35b9a2508" />


*   **➕ 過濾器建議設定:**
    *   這個設定可以確保顯示出還沒驗證完全的弱點。\
        <img width="1092" alt="image" src="https://github.com/user-attachments/assets/60f075d5-4971-49bd-9f0b-232574aefa8c" />

*   **➕ 如何開始:**
    *   建立專案資料夾
    *   放入 .xml 掃描結果
    *   放入 .scan 掃描檔
    *   開啟應用程式
    *   使用瀏覽器開啟 Web 應用
