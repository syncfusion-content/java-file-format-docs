---
title: Working with Document Conversions
description: This section illustrates how to convert a Word document into other supported file formats using Syncfusion Java Word library
platform: java-file-formats
control: Word Library
documentation: UG
---

# Word Document Conversion

## Working with Document Conversions

The Essential DocIO converts documents from one format to another format. Each file format document can be categorized as flow layout document or fixed layout document.

**Flow layout document**

* A flow document is designed to "reflow content" depending on the application.
* Does not contain any information about the position of its content.
* Dynamically renders the content by application at run time.
* Example: DOCX, HTML and TEXT file formats.

Essential DocIO can convert various flow document as fixed document by using our layout engine. Following conversions are supported by Essential DocIO.

* Microsoft Word file format Conversions.
* Text Conversions.
* HTML Conversions.

## HTML conversion

Essential DocIO supports converting the HTML file into Word document and vice versa. It supports only the HTML files that meet the validation either against XHTML 1.0 strict or XHTML 1.0 Transitional schema. 

For further information kindly refer here.

### Customizing the HTML to Word conversion

You can customize the HTML to Word conversion with the following options:

* Validate the HTML string against XHTML 1.0 Strict and Transitional schema
* Insert the HTML string at the specified position of the document body contents
* Append HTML string to the specified paragraph

For further information kindly refer this link.

### Customizing the Word to HTML conversion

You can customize the Word to HTML conversion with the following options:

* Extract the images used in the HTML document at the specified file directory 
* Specify to export the header and footer of the Word document in the HTML 
* Specify to consider Text Input field as a editable fields or text 
* Specify the CSS style sheet type and its name

N> 
While exporting header and footer, DocIO exports the first section header content at the top of the HTML file and first section footer content at the end of the HTML file

For further information kindly refer this link.

### Supported Document elements

Kindly refer to this link for the document elements and attributes that are supported by DocIO in the Word to HTML and HTML to Word conversions.

## Text file

Essential DocIO supports to convert the Word document into a Text file and vice versa. For further information, kindly refer to this link.