---
title: Word Document to HTML conversion | Word library | Syncfusion
description: This section illustrates how to convert Word document to HTML using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Word to HTML and HTML to Word Conversions

The Essential DocIO converts the HTML file into a Word document and vice versa. You can also convert the Word document (DOCX, DOTX, DOCM, and DOTM) into HTML format.

In the Word library (DocIO) we use the [XmlReader](https://docs.microsoft.com/en-us/dotnet/api/system.xml.xmlreader?view=netframework-4.8) for parsing the content from input HTML. So, the input HTML should meet the XML standard (have proper open and close tags), even if you specify the `XHTMLValidationType` parameter as `XHTMLValidationType.None`.

## XHTML Validation

Every HTML content is validated against a Document Type Declaration (DTD) which is a set of mark-up declarations that define a document type for a SGML-family mark-up language (GML, SGML, XML, HTML).

### XHTML validation types

The following XHTML validation types are supported in Essential DocIO while importing an HTML content.

<table>
<thead>
<tr>
<td>XHTML validation types</td>
<td>Description</td>
</tr>
</thead>
<tr>
<td><b>XHTMLValidationType.None</b></td>
<td>It does not perform any schema validation but the given HTML content should meet the XHTML 1.0 format.</td>
</tr>
<tr>
<td><b>XHTMLValidationType.Transitional</b></td>
<td>It allows several attributes within the tags.</td>
</tr>
<tr>
<td><b>XHTMLValidationType.Strict</b></td>
<td>It does not allow the attributes inside the tag.</td>
</tr>
</table>

The following code example shows how to convert the HTML file into a Word document.

{% tabs %}
{% highlight JAVA %}
//Load the HTML document against the validation type none.
WordDocument document = new WordDocument("Input.html", FormatType.Html);
document.setXHTMLValidateOption(XHTMLValidationType.None);
//Save the Word document.
document.save("HTMLtoWord.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

The following code example shows how to convert the Word document into HTML.

{% tabs %}
{% highlight JAVA %}
//Load the template document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Saves the document as an Html file.
document.save("WordToHtml.html", FormatType.Html);
//Close the document. 
document.close();
{% endhighlight %}
{% endtabs %}			

## Supported and unsupported items

The following document elements and attributes are supported by DocIO in Word to HTML and HTML to Word conversions.

<table>
<thead> 
<tr>
<th>Document Element</th>
<th>Attribute</th>
<th>Support Status</th>
<th>Notes</th>
</tr>
</thead>
<tr>
<td>
Bookmark<br/><br/></td>
<td>
Id<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Border<br/><br/><br/><br/></td>
<td>
Color<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Distance from text<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Line style<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Some line styles are rendered as solid.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Line width<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Document Properties<br/><br/></td>
<td>
<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Field<br/><br/></td>
<td>
<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Footnotes and Endnotes<br/><br/></td>
<td>
<br/><br/></td>
<td>
No<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Form Field<br/><br/></td>
<td>
Text input, Checkbox and combo box<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Header / Footer<br/><br/></td>
<td>
Different per section<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Only odd header of the first section is preserved in HTML export.<br/><br/></td>
</tr>
<tr>
<td>
Hyperlink<br/><br/></td>
<td>
External URL<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Local<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Image<br/><br/></td>
<td>
Inline<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Scale<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
List<br/><br/></td>
<td>
Custom bullets<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Multi-level<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Numbered<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Restart numbering<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Standard bullets<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Comment<br/><br/></td>
<td>
<br/><br/></td>
<td>
No<br/><br/></td>
<td>
<br/><br/></td>
</tr>
<tr>
<td>
Symbols<br/><br/></td>
<td>
<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
<br/><br/></td>
</tr>
<tr>
<td>
Paragraph<br/><br/></td>
<td>
Alignment<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Borders<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
See Borders, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Keep lines and paragraphs together<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Paragraph Indents<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Line spacing<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Page break before<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Shading<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
See Shading, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Spacing before and after<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Shading<br/><br/><br/><br/></td>
<td>
Background color<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Solid background colors are supported.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Foreground color<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Solid foreground color is used when background color is auto.<br/><br/></td>
</tr>
<tr>
<td>
Styles<br/><br/><br/><br/></td>
<td>
Paragraph styles<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Character styles<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
List styles<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Table<br/><br/><br/><br/></td>
<td>
Alignment<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Cell margins<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Column widths<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Indent from left<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Preferred width<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Spacing between cells<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Borders<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
See Borders, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Shading<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
See Shading, for more details.<br/><br/></td>
</tr>
<tr>
<td>
Nested Table<br/><br/></td>
<td>
<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
<br/><br/></td>
</tr>
<tr>
<td>
Table Cell<br/><br/><br/><br/></td>
<td>
Borders<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
See Borders, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Cell margins<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Horizontal merge<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Shading<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
See Shading, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Vertical alignment<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Vertical merge<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Table Row<br/><br/></td>
<td>
Height<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Padding<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
Text<br/><br/><br/><br/></td>
<td>
All caps<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Bold<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Character spacing<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Color<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Emboss<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Rendered as bold.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Engrave<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Rendered as bold.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Font<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Hidden<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Highlighting<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Imprint<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Rendered as bold.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Italic<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Line breaks<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Outline<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Rendered as bold.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Page breaks<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Shading<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
See Shading, for more details.<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Small caps<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Special symbols<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Strike out<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Subscript / Superscript<br/><br/></td>
<td>
Yes<br/><br/></td>
<td>
-<br/><br/></td>
</tr>
<tr>
<td>
<br/><br/></td>
<td>
Underline<br/><br/></td>
<td>
Partial<br/><br/></td>
<td>
Underline types and colors are ignored.
</td>
</tr>
</table>
