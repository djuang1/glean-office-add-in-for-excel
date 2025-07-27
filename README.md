# Glean Office Add-in for Excel

## Summary

Glean Office Add-In for Excel. This add-in allows you to quickly search your Glean instance from Excel. It provides a custom function that will perform a search against a cell or a text string.

## Features

- Glean inside Excel

## Applies to

- Excel on Windows, Mac, and in a browser.

## Prerequisites

- Microsoft Office 365 

### Manifest

The manifest file is an XML file that describes an add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly the manifest provides URL locations for where Office can find and download the add-in's resource files.

- **manifest.xml**: This manifest file gets the add-in's HTML page from this repo. This is the quickest way to try out the sample. 

## Run the sample on Excel on web

The Glean Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. 

The Glean Office Add-in for Excel is configured so that the files are hosted directly from this GitHub repo. Use the following steps to sideload the manifest.xml file to see the sample run.

1.  Download the **manifest.xml** file from the sample folder for Excel.
1.  Open [Office on the web](https://office.live.com/).
1.  Choose **Excel**, and then open a new document.
1.  On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.
1.  On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
2.  Browse to the add-in manifest file, and then select **Upload**.
3.  Verify that the add-in loaded successfully. You will see a **Glean** button on the **Home** tab on the ribbon.

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/djuang1/glean-office-add-in-for-excel/issues/new/choose) and we'll help you out.

## Copyright

Copyright (c) 2025 dejim.com
