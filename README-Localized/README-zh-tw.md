# 影像圖說文字 Word 增益集範例︰載入、編輯和插入影像

**目錄**

* [摘要](#summary)
* [需要的工具](#required-tools)
* [如何安裝憑證](#how-to-install-certificates)
* [如何設定和執行應用程式](#how-to-set-up-and-run-the-app)
* [如何在 Word 2016 for Windows 中執行增益集](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [常見問題集](#faq)
* [問題和建議](#questions-and-comments)
* [深入了解](#learn-more)


## 摘要

此 Word 增益集範例將會為您示範如何：

1. 使用 Typescript，以建立 Word 增益集。
2. 從文件將影像載入增益集。
3. 在增益集中編輯影像，方法是使用 HTML 畫布 API 並且將影像插入 Word 文件。
4. 實作增益集命令，從功能區啟動增益集，以及從功能區與內容功能表執行指令碼。
5. 使用 Office UI Fabric，建立增益集的原生類似 Word 的經驗。

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

定義 - **增益集命令**：Word UI 的擴充，可讓您在工作窗格中啟動增益集，或從功能區或內容功能表執行指令碼。

如果您只想要看到此動作，直接跳到 [Word 2016 for Windows 設定](#word-2016-for-windows-set-up)，並使用這個[資訊清單](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml)。

## 需要的工具

若要使用影像圖說文字 Word 增益集範例，需要有下列各項。

* Word 2016 16.0.6326.0000 或更高版本，或支援 Word Javascript API 的任何用戶端。這個範例會執行必要檢查，以查看它是否在 JavaScript API 支援的主機中執行。
* npm (https://www.npmjs.com/) 以安裝相依性。它隨附 [NodeJS](https://nodejs.org/en/)。
* (Windows) [Git Bash](http://www.git-scm.com/downloads).
* 複製此儲存機制到本機電腦。

> 附註：Word for Mac 2016 目前不支援增益集命令。這個範例可以在 Mac 上執行，不需要增益集命令。

## 如何安裝憑證

您需要憑證才能執行這個範例，因為增益集命令需要 HTTPS 且因為增益集命令無 UI，您無法接受無效的憑證。執行 [./gen-cert.sh](#gen-cert.sh) 來建立憑證，然後您必須將 ca.crt 安裝到您的信任根目錄憑證授權存放區 (Windows)。

## 如何設定和執行應用程式

1. 安裝 TypeScript 定義管理員，方法是在命令列輸入 ```npm install typings -g```。
2. 安裝 typings.json 中識別的 Typescript 定義，方法是在專案根目錄的命令列執行 ```typings install```。
3. 安裝 package.json 中識別的專案相依性，方法是在專案的根目錄執行 ```npm install```。
4. 安裝 gulp ```npm install -g gulp```。
5. 複製 Fabric 和 JQuery 檔案，方法是執行 ```gulp copy:libs```。(Windows) 如果您在這裡有問題，請確定 *%APPDATA%\npm* 在您的路徑變數中。
6. 執行預設 gulp 工作，方法是從專案的根目錄執行 ```gulp```。如果 TypeScript 定義沒有更新，您會在這裡遇到錯誤。

目前您已部署這個範例增益集。現在，您需要讓 Word 知道哪裡可以找到此增益集。

### Word 2016 for Windows 設定

1. (僅限 Windows) 解壓縮及執行這個[登錄機碼](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark)來啟動增益集命令功能。當增益集命令是**預覽功能**時，這是必要項目。
2. 建立網路共用，或[共用網路資料夾](https://technet.microsoft.com/zh-tw/library/cc770880.aspx)，並將 [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) 資訊清單檔案放置在其中。
3. 啟動 Word 並開啟一個文件。
4. 選擇 [檔案]<e /> 索引標籤，然後選擇 [選項]<e />。
5. 選擇 [信任中心]<e />，然後選擇 [信任中心設定]<e /> 按鈕。
6. 選擇 [受信任的增益集目錄]<e />。
7. 在 [目錄 URL]<e /> 方塊中，輸入包含 manifest-word-add-in-canvas.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]<e />。
8. 選取 [顯示於功能表中]<e /> 核取方塊，然後選擇 [確定]<e />。
9. 接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。關閉並重新啟動 Word。

## 如何在 Word 2016 for Windows 中執行增益集

1. 開啟 Word 文件。
2. 在 Word 2016 的 [插入]<e /> 索引標籤上，選擇 [我的增益集]<e />。
3. 選取 [共用資料夾]<e /> 索引標籤。
4. 選擇 [影像圖說文字增益集]<e />，然後選取 [插入]<e />。
5. 如果您的 Word 版本支援增益集命令，UI 會通知您已載入增益集。您可以使用 [圖說文字增益集]<e /> 索引標籤，在 UI 中載入增益集，並將影像插入文件。您也可以使用右鍵內容功能表來將影像插入文件。
6. 如果您的 Word 版本不支援增益集命令，增益集會載入工作窗格。您必須將圖片插入 Word 文件，以使用增益集的功能。
7. 在 Word 文件中選取影像，將其載入工作窗格，方法是從 doc* 選取*載入影像。您現在可以將圖說文字插入影像。選取 *將影像插入文件* 已將更新的影像插入 Word 文件。增益集將會產生每個圖說文字的預留位置描述。

## 常見問題集

* 增益集命令是否可在 Mac 和 iPad 上運作？否，在此讀我檔案發佈時，無法在 Mac 或 iPad 上運作。
* 為什麼我的增益集未顯示在 [我的增益集]<e /> 視窗中？您的增益集資訊清單可能有錯誤。建議您針對[資訊清單結構描述](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD)驗證資訊清單。
* 為什麼無法針對我的增益集命令呼叫函數檔案？增益集命令需要 HTTPS。由於增益集命令需要 TLS，而且沒有 UI，您無法查看是否有憑證問題。如果您必須在工作窗格中接受不正確的憑證，則增益集命令無法運作。
* 為什麼 npm 安裝命令懸置？可能不是懸置。在 Windows 上需要一段時間。

## 問題和建議

我們很樂於收到您對於影像圖說文字 Word 增益集範例的意見反應。您可以在此儲存機制的[問題](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues)區段中，將您的問題及建議傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關增益集開發的一般問題。務必以 [office-js]、[word-addins] 和 [API] 標記您的問題或意見。我們會查看這些標記。

## 深入了解

這裡有更多的資源，可協助您建立以 Word Javascript API 為基礎的增益集︰

* [Office 增益集平台概觀](https://msdn.microsoft.com/zh-tw/library/office/jj220082.aspx)
* [Word 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Word 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)
* [SillyStories 範例](https://github.com/OfficeDev/Word-Add-in-SillyStories) - 了解如何從服務中載入 docx 檔案，以及將檔案插入開啟的 Word 文件。
* [Node.js 的 Office 增益集伺服器驗證範例](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) - 了解如何使用 Azure 和 Google OAuth 提供者來驗證增益集使用者。

## 著作權
Copyright (c) 2016 Microsoft.著作權所有，並保留一切權利。
