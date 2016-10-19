# <a name="image-callouts-word-add-in-sample:-load,-edit,-and-insert-images"></a>图像标注 Word 外接程序示例：加载、编辑和插入图像

**目录**

* [摘要](#summary)
* [必备工具](#required-tools)
* [如何安装证书](#how-to-install-certificates)
* [如何设置并运行应用](#how-to-set-up-and-run-the-app)
* [如何在 Word 2016 for Windows 中运行外接程序](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [常见问题解答](#faq)
* [问题和意见](#questions-and-comments)
* [了解更多](#learn-more)


## <a name="summary"></a>摘要

此 Word 外接程序示例演示如何：

1. 使用 Typescript 创建 Word 外接程序。
2. 将文档中的图像加载到外接程序中。
3. 通过使用 HTML 画布 API 编辑外接程序中的图像并将图像插入到 Word 文档中。
4. 实现外接程序命令，从功能区启动外接程序并从功能区和上下文菜单中运行脚本。
5. 使用 Office UI 结构为你的外接程序创建本机 Word 般的体验。

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

定义 - **外接程序命令**：Word UI 的扩展，允许你在任务窗格中启动外接程序，或者从功能区或上下文菜单中运行脚本。

如果你只想看到实际操作，请跳到 [Word 2016 for Windows 设置](#word-2016-for-windows-set-up)，并使用此[清单](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml)。

## <a name="required-tools"></a>必需的工具

若要使用图像标注 Word 外接程序示例，必须符合以下条件。

* Word 2016 16.0.6326.0000 或更高版本，或任何支持 Word Javascript API 的客户端。此示例会执行要求检查以查看是否正在受支持的 JavaScript API 主机中运行。
* 可安装依赖项的 npm (https://www.npmjs.com/)。它附带了 [NodeJS](https://nodejs.org/en/)。
* (Windows) [Git Bash](http://www.git-scm.com/downloads)。
* 克隆此存储库到本地计算机。

> 注意：目前，Word for Mac 2016 不支持外接程序命令。此示例无需外接程序命令即可在 Mac 上运行。

## <a name="how-to-install-certificates"></a>如何安装证书

你必须有证书才能运行此示例，因为外接程序命令需要 HTTPS，而且由于外接程序命令无 UI，你不能接受无效的证书。运行 [./gen-cert.sh](#gen-cert.sh) 创建证书，然后你需要将 ca.crt 安装到受信任的根证书颁发机构存储区中 (Windows)。

## <a name="how-to-set-up-and-run-the-app"></a>如何设置并运行应用

1. 通过在命令行处键入 ```npm install typings -g``` 来安装 TypeScript 定义管理器。
2. 通过在命令行处运行项目的根目录中的 ```typings install``` 来安装在 typings.json 中标识的 Typescript 定义。
3. 通过在项目的根目录中运行 ```npm install``` 来安装在 package.json 中标识的项目依赖项。
4. 安装 gulp ```npm install -g gulp```。
5. 通过运行 ```gulp copy:libs``` 复制结构和 JQuery 文件。(Windows) 如果你在此处遇到问题，请确保 *%APPDATA%\npm* 位于你的 path 变量中。
6. 通过从项目的根目录运行 ```gulp``` 来运行默认的 gulp 任务。如果没有更新 TypeScript 定义，你将在此处遇到错误。

此时，你已部署了第一个示例外接程序。现在，你需要让 Word 知道在哪里可以找到该外接程序。

### <a name="word-2016-for-windows-set-up"></a>Word 2016 for Windows 设置

1. （仅限 Windows）解压缩并运行此[注册表项](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark)以激活外接程序命令功能。当外接程序命令是**预览功能**时必须执行此操作。
2. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/en-us/library/cc770880.aspx)，并将 [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) 清单文件放入该文件夹中。
3. 启动 Word，然后打开一个文档。
4. 选择**文件**选项卡，然后选择**选项**。
5. 选择**信任中心**，然后选择**信任中心设置**按钮。
6. 选择“**受信任的外接程序目录**”。
7. 在“**目录 URL**”字段中，输入包含 manifest-word-add-in-canvas.xml 的文件夹共享的网络路径，然后选择“**添加目录**”。
8. 选择“**显示在菜单中**”复选框，然后选择“**确定**”。
9. 随后会出现一条消息，告知您下次启动 Office 时将应用您的设置。关闭并重新启动 Word。

## <a name="how-to-run-the-add-in-in-word-2016-for-windows"></a>如何在 Word 2016 for Windows 中运行外接程序

1. 打开一个 Word 文档。
2. 在 Word 2016 中的**插入**选项卡上，选择**我的外接程序**。
3. 选择**共享文件夹**选项卡。
4. 选择“**图像标注外接程序**”，然后选择“**插入**”。
5. 如果你的 Word 版本支持外接程序命令，UI 将通知你加载了外接程序。可以使用“**标注外接程序**”选项卡在 UI 中加载外接程序以及将图像插入文档中。还可以使用右键单击上下文菜单将图像插入到文档中。
6. 如果你的 Word 版本不支持外接程序命令，则外接程序将在任务窗格中加载。需要将图片插入到 Word 文档中才能使用外接程序的功能。
7. 在 Word 文档中选择一个图像，并通过选择“*从 doc 加载图像*”将其加载到任务窗格中。现在，你可以将标注插入到图像中。选择“*将图像插入到 doc 中*”将更新的图像放入 Word 文档中。外接程序将针对各个标注生成相应的占位符说明。

## <a name="faq"></a>常见问题解答

* 外接程序命令将在 Mac 和 iPad 上可用？不，截至此自述文件发布之前，这些命令都将无法在 Mac 或 iPad 上使用。
* 为什么我的外接程序显示在“**我的外接程序**”窗口中？你的外接程序清单可能有错误。我建议你针对[清单架构](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD)验证清单。
* 为什么我的外接程序命令不调用该函数文件？外接程序命令需要 HTTPS。由于外接程序命令需要 TLS，并且没有 UI，因此你无法看到是否存在证书问题。如果你必须在任务窗格中接受无效的证书，那么该外接程序命令不起作用。
* 为什么 npm 安装命令挂起？很可能不是挂起，只是在 Windows 上需要一段时间。

## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对图像标注 Word 外接程序示例的相关反馈。你可以在该存储库中的“[问题](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues)”部分将问题和建议发送给我们。

与外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)。确保你的问题或意见使用了 [office-js]、[word-addins] 和 [API] 标记。我们会一直关注这些标记。

## <a name="learn-more"></a>了解详细信息

以下是更多的资源，可帮助你创建基于 Word Javascript API 的外接程序：

* [Word 外接程序文档和示例](https://dev.office.com/word)
* [SillyStories 示例](https://github.com/OfficeDev/Word-Add-in-SillyStories) - 了解如何从服务中加载 docx 文件以及将文件插入到打开的 Word 文档中。
* [Node.js 的 Office 外接程序服务器身份验证示例](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) - 了解如何使用 Azure 和 Google OAuth 提供程序对外接程序用户进行身份验证。

## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft。保留所有权利。
