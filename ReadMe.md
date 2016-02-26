# Image callouts Word add-in sample: load, edit, and insert images

**Table of contents**

* [Summary](#summary)
* [Required tools](#required-tools)
* [How to install certificates](#how-to-install-certificates)
* [How to setup and run the app](#how-to-setup-and-run-the-app)
* [How to run the add-in in Word 2016 for Windows](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [FAQ](#faq)
* [Questions and comments](#questions-and-comments)
* [Learn more](#learn-more)


## Summary

This Word add-in sample shows you how to:

1. Create a Word add-in with Typescript.
2. Load images from the document into the add-in.
3. Edit images in the add-in by using the HTML canvas API and insert the images into a Word document.
4. Implement add-in commands that both launch an add-in from the ribbon and run a script from both the ribbon and a context menu.
5. Use the Office UI Fabric to create a native Word-like experience for your add-in.

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

Definition- **add-in command**: an extension to the Word UI that allows you to either launch the add-in in a task pane or run a script, from either the ribbon or a context menu.

If you just want to see this in action, skip to [Word 2016 for Windows setup](#word-2016-for-windows-setup) and use this [manifest](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml).

## Required tools

To use the Image callouts Word add-in sample, the following are required.

* Word 2016 16.0.6326.0000 or higher, or any client that supports the Word Javascript API. This sample does a requirement check to see if it is running in a supported host for the JavaScript APIs.
* npm (https://www.npmjs.com/) to install the dependencies. It comes with [NodeJS](https://nodejs.org/en/).
* (Windows) [Git Bash](http://www.git-scm.com/downloads).
* Clone this repo to your local computer.

> Note: Word for Mac 2016 does not support add-in commands at this time. This sample can run on the Mac without the add-in commands.

## How to install certificates

You'll need a certificate to run this sample since add-in commands require HTTPS and since add-in commands are UI-less, you can't accept invalid certificates. Run [./gen-cert.sh](#gen-cert.sh) to create the certificate and then you'll need to install ca.crt into your Trusted Root Certification Authorities store (Windows).

## How to setup and run the app

1. Install the TypeScript definition manager by typing ```npm install typings -g``` at the command line.
2. Install the Typescript definitions identified in typings.json by running ```typings install``` in the project's root directory at the command line. Note that the TypeScript definitions are out of date and will cause errors. You'll need to fake the missing definitions until the official definitions are updated on DefinatelyTyped. The definitions are in a directory called typings.
3. Install the project dependencies identified in package.json by running ```npm install``` in the project's root directory.
4. Install gulp ```npm install -g gulp```.
5. Copy the Fabric and JQuery files by running ```gulp copy:libs```. (Windows) If you have an issue here, make sure that *%APPDATA%\npm* is in your path variable.
6. Run the default gulp task by running ```gulp``` from the project's root directory. If the TypeScript definitions aren't updated, you'll get an error here.

You've deployed this sample add-in at this point. Now you need to let Word know where to find the add-in.

### Word 2016 for Windows setup

1. (Windows only) Unzip and run this [registry key](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark) to activate the add-in commands feature. This is required while add-in commands are a **preview feature**.
2. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx) and place the [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) manifest file in it.
3. Launch Word and open a document.
4. Choose the **File** tab, and then choose **Options**.
5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
6. Choose **Trusted Add-ins Catalogs**.
7. In the **Catalog Url** box, enter the network path to the folder share that contains manifest-word-add-in-canvas.xml and then choose **Add Catalog**.
8. Select the **Show in Menu** check box, and then choose **OK**.
9. A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Word.

## How to run the add-in in Word 2016 for Windows

1. Open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **Shared folder** tab.
4. Choose **Image callout add-in**, and then select **Insert**.
5. If add-in commands are supported by your version of Word, the UI will inform you that the add-in was loaded. You can use the **Callout add-in** tab to load the add-in in the UI and to insert an image into the document. You can also use the right-click context menu to insert an image into the document.
6. If add-in commands are not supported by your version of Word, the add-in will load in a task pane. You'll need to insert a picture into the Word document to use the functionality of the add-in.
7. Select an image in the Word document, and load it into the taskpane by selecting *Load image from doc*. You can now insert callouts into the image. Select *Insert image into doc* to place the updated image into the Word doc. The add-in wil generate placeholder descriptions for each of the callouts.

## FAQ

* Will add-in commands work on Mac and iPad? No, they won't work on the Mac or iPad as of the publication of this readme.
* Why doesn't my add-in show up in the **My Add-ins** window? Your add-in manifest may have an error. I suggest that you validate the manifest against the [manifest schema](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD).
* Why doesn't the function file get called for my add-in commands? Add-ins commands require HTTPS. Since the add-in commands require TLS, and there isn't a UI, you can't see whether there is a certificate issue. If you have to accept an invalid certificate in the taskpane, then the add-in command will not work.
* Why do npm install commands hang? It probably isn't hung. It just takes a while on Windows.

## Questions and comments

We'd love to get your feedback about the Image callout Word add-in sample. You can send your questions and suggestions to us in the [issues](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues) section of this repository.

Questions about add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [office-js], [word-addins], and [API]. We are watching these tags.

## Learn more

Here are more resources to help you create Word Javascript API based add-ins:

* [Office Add-ins platform overview](https://msdn.microsoft.com/EN-US/library/office/jj220082.aspx)
* [Word add-ins](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)
* [SillyStories sample](https://github.com/OfficeDev/Word-Add-in-SillyStories) - learn how to load docx files from a service and insert the files into an open Word document.
* [Office Add-in Server Authentication Sample for Node.js](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) - learn how use Azure and Google OAuth providers for authenticating add-in users.

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.