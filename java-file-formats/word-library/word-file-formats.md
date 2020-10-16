---
title: Word file format conversions | Word library | Syncfusion
description: This section illustrates Word file format conversions supported in Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---


# Word File Formats in Essential DocIO

The [Microsoft Word's](https://en.wikipedia.org/wiki/Microsoft_Word#) native file formats are DOCX, DOTX, DOCM, and DOTM. The Essential DocIO supports the following major native file formats.

1. Word Open XML formats (2007 & later)
2. Word Processing XML (.xml)

## Word Open XML formats (2007 & later)

[Office Open XML](http://en.wikipedia.org/wiki/Office_Open_XML#) (OOXML or Microsoft Open XML (MOX)) is a zipped, new XML-based file format introduced by Microsoft in Office 2007 applications.The WordprocessingML is the markup language used by the Microsoft Office Word to store its DOCX documents.

DocIO supports the following WordprocessingML:

* Microsoft Word 2007
* Microsoft Word 2010
* Microsoft Word 2013
* Microsoft Word 2016
* Microsoft Word 2019

The following code example explains how to create a new Word document with a few lines of code.

{% tabs %}
{% highlight JAVA %}
//Create an instance of the WordDocument Instance (Empty Word Document).
WordDocument document = new WordDocument();
//Add a section and a paragraph in the empty document.
document.ensureMinimal();
//Append text to the last paragraph of the document.
document.getLastParagraph().appendText("Hello World");
//Save and close the Word document.
document.save("Sample.docx");
document.close();
{% endhighlight %}

{% endtabs %}

### Templates

DOTX is a Word document template. The following code sample shows how to create the Word document template with a few lines of code.

{% tabs %}
{% highlight JAVA %}
//Create an instance of the WordDocument Instance (Empty Word Document).
WordDocument document = new WordDocument();
//Add a section and a paragraph in the empty document.
document.ensureMinimal();
//Append text to the last paragraph of the document.
document.getLastParagraph().appendText("Hello World");
//Save and close the Word document.
document.save("Sample.dotx");
document.close();

{% endhighlight %}
{% endtabs %}

### Macros

DOCM is a macro-enabled Word document. It is same as the DOCX document contains macros and scripts. The DocIO provides only preservation support for macros. The following code shows how to load and save a macro-enabled document using the DocIO library.

{% tabs %}
{% highlight JAVA %}
// Load the macro-enabled template.
WordDocument document = new WordDocument("Template.dotm");
// Get the table.
DataTableSupport table = getDataTable();
// Execute the Mail Mmrge with groups.
document.getMailMerge().executeGroup(table);
//Save and close the document.
document.save("Sample.docm", FormatType.Word2013Docm);
document.close();
{% endhighlight %}
{% endtabs %}

## Word Processing XML (.xml)

The XML format introduced in Microsoft Word 2003 was a simple, XML-based format called WordprocessingML or WordML.
The Essential DocIO supports converting the Word document into Word Processing XML document and vice versa.

N> 1. Importing and exporting the Word Processing 2007 XML documents is supported.
N> 2. Exporting the Word Processing 2003 XML document is not supported. Whereas you can import the Word Processing 2003 XML documents and export it to the other supported file formats.
N> 3. The custom XML elements present in the Word Processing 2003 XML documents will be removed automatically while importing, like latest Microsoft Word. The custom XML element is a depreciated feature in latest Microsoft Word.

The following code example shows how to convert the Word document into Word Processing XML document.

{% tabs %}
{% highlight JAVA %}
//Load an existing Word document.
WordDocument document = new WordDocument("Sample.docx");
//Save the document as a Word Processing ML document.
document.save("WordToWordML.xml", FormatType.WordML);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

The following code example shows how to convert the Word Processing XML document into Word document.

{% tabs %}
{% highlight JAVA %}
// Load an existing Word document. 
WordDocument document = new WordDocument("Template.xml");
//Save the Word Processing ML document as docx.
document.save("WordMLToWord.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

### Unsupported elements in Word to Word Processing XML conversion:

The following table contains a list of unsupported elements in the Word to Word Processing XML conversion.

<table>
<thead> 
<tr>
<th>Element</th>
<th>Limitations or Unsupported elements</th>
</tr>
</thead>
<tr>
<td>
Custom Shapes<br/><br/></td>
<td>
Not supported<br/><br/></td>
</tr>
<tr>
<td>
Embedded fonts<br/><br/></td>
<td>
Not supported<br/><br/></td>
</tr>
<tr>
<td>
Equation<br/><br/></td>
<td>
Not supported<br/><br/></td>
</tr>
<tr>
<td>
SmartArt<br/><br/></td>
<td>
Not supported<br/><br/></td>
</tr>
<tr>
<td>
WordArt<br/><br/></td>
<td>
Not supported<br/><br/></td>
</tr>
<tr>
<td>
Form Fields
</td>
<td>
Unparsed in Word Processing 2003 XML document
</td>
</tr>
<tr>
<td>
Ole Object
</td>
<td>
Not supported
</td>
</tr>
</table>
