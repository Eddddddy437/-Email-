# -Email-
加班Emial申請UI+自動偵測回信並下載回信
# 📧 Overtime Application & Response Monitor
### 辦公自動化：加班申請自動發信與回信監控存檔工具

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg) ![Library](https://img.shields.io/badge/Library-pywin32-orange.svg) ![GUI](https://img.shields.io/badge/GUI-Tkinter-green.svg)

## 📖 專案簡介
本專案是一個基於 Python 開發的辦公輔助工具，旨在簡化企業內部的加班申請流程。透過整合 **Tkinter (GUI)** 與 **Microsoft Outlook (Win32COM)**，實現一鍵生成申請郵件、自動統計月度加班頻率，並在背景實時監控主管回信，達成自動化存檔。

## ✨ 核心功能
* **自動化郵件生成**：自定義加班日期、時間與事由，一鍵調用 Outlook 視窗並自動填入收件人、副本及簽名檔。
* **月度統計戰報**：自動掃描 Outlook「已傳送郵件」，統計並顯示本月已申請加班的次數與日期，並附上人性化的激勵小語。
* **背景監控回信 (Threaded Monitor)**：使用背景執行緒監控收件匣，當偵測到主管的 RE (回覆) 信件時，自動將該郵件匯出為 `.msg` 檔並存至桌面。
* **檔案去標識化**：自動處理檔名中的非法字元，確保存檔路徑穩定且相容於不同電腦環境。

## 🛠️ 技術棧
* **語言**：Python 3
* **GUI 介面**：`Tkinter`
* **Office 整合**：`pywin32` (Win32COM 接口)
* **多執行緒**：`threading` & `pythoncom` (用於非阻塞式背景監控)
* **正則表達式**：`re` (用於自動解析郵件標題日期)

## 🚀 如何使用
1. **環境準備**：
   確保您的電腦已安裝 Microsoft Outlook 並登入帳號。
2. **安裝依賴套件**：
   ```bash
   pip install pywin32
