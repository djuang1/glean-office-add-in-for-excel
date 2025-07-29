# Glean Office Add-in for Excel

## Summary

![Screenshot of Glean Office Add-in for Excel.](https://raw.githubusercontent.com/djuang1/glean-office-add-in-for-excel/refs/heads/main/assets/screenshot.png)

Glean Office Add-In for Excel. This proof-of-concept add-in allows you to quickly search your Glean instance from Excel. It provides a custom function that will perform a search against a cell or a text string.

> [!NOTE]  
This Office add-in currently requires the use of a Chrome plugin to allow CORS since the add-in is hosted directly from this Github repo.

## Applies to

- Excel on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft Office 365 - Excel
- [Chrome Extension - Allow CORS](https://chromewebstore.google.com/detail/allow-cors-access-control/lhobafahddgcelffkeicbaginigeejlf)

## Setup and Run

Since this a proof-of-concept currently, the Glean Office Add-in for Excel is configured so that the files are hosted directly from this GitHub repo. Use the following steps to sideload the manifest.xml file to use the add-in in Excel.

1.  Download the **manifest.xml** file 
2.  Open [Office on the web](https://office.live.com/).
3.  Choose **Excel**, and then open a new document.
4.  On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.
5.  On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
6.  Browse to the add-in manifest file, and then select **Upload**.
7.  Verify that the add-in loaded successfully. You will see a **Glean** button on the **Home** tab on the ribbon.
8.  Click to open the task pane. Enter in your Glean instance name and the Client API token. The client API token should have the **Search** scope enabled.
9.  Click on **Save**
10. In any cell, type in ```=Glean.Search("Hello")``` and you'll see the add-in go and grab the answer from Glean and fill in the cell.

## Questions and feedback

- Did you experience any problems with the Office Add-in? [Create an issue](https://github.com/djuang1/glean-office-add-in-for-excel/issues/new/choose) and we'll help you out.

## Copyright

Copyright (c) 2025 dejim.com
