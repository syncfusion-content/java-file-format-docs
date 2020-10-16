---
title: FAQ Section | Word library | Syncfusion
description: This section illustrates about Frequently Asked Questions in Essential Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Frequently Asked Questions

The frequently asked questions in Essential DocIO are listed below.

## How to modify an existing style?

The following code illustrates how to modify the built-in style while creating new Word document.

{% tabs %}   

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Create built-in style and modifies its properties
Style style = (Style)Style.createBuiltinStyle(BuiltinStyle.Heading1, document);
style.getCharacterFormat().setItalic(true);
style.getCharacterFormat().setTextColor(ColorSupport.getDarkGreen());
//Add it to the styles collection.
document.getStyles().add(style);
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
IWTextRange text = paragraph.appendText("A built-in style is modified and is applied to this paragraph.");
//Apply the new style to paragraph.
paragraph.applyStyle(style.getName());
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}  

{% endtabs %} 

## How to insert a table from HTML string in Word document?

An HTML string can be inserted to the Word document at text body or paragraph. The following code illustrates how to insert a table to the document from the HTML string.

{% tabs %}  

{% highlight JAVA %}
//Loads the template document
WordDocument document = new WordDocument("Template.docx");
//Gets the text body
WTextBody textbody = document.getSections().get(0).getBody();
//Html string that represents table with two rows and two columns
String htmlString = " <table border='1'><tr><td><p>First Row First Cell</p></td><td><p>First Row Second Cell</p></td></tr><tr><td><p>Second Row First Cell</p></td><td><p>Second Row Second Cell</p></td></tr></table> ";
//Inserts the string to the text body
textbody.insertXHTML(htmlString);
//Saves and closes the document
document.save("Sample.docx");
document.close();

{% endhighlight %}

 {% endtabs %}  

 
## How to set table cell width?

Each cell in the table can have its own width. The following code illustrates how to set the width of the cell.

{% tabs %}  

{% highlight JAVA %}
// Open word document.
WordDocument document = new WordDocument("Template.docx");
// Get the text body of first section.
WTextBody textbody = document.getSections().get(0).getBody();
// Get the table.
IWTable table = textbody.getTables().get(0);
// Iterate through table rows.
for (Object rows_tempObj : table.getRows()) 
{
    WTableRow row = (WTableRow) rows_tempObj;
// Set width for cells.
for (int i = 0; i < row.getCells().getCount(); i++) 
{
     WTableCell cell = row.getCells().get(i);
     if (i % 2 == 0)
     // Set width as 100 for cells in even column.
        cell.setWidth(100);
     else
     // Set width as 150 for cell in odd column.
        cell.setWidth(150);
}
}
// Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## How to position a table in a Word document?

You can position a table in a Word document by setting position properties. The following code illustrates how to set position properties for a table.

{% tabs %}  

{% highlight JAVA %}
//Load the template document.
WordDocument document = new WordDocument("Template.docx");
//Get the text body of first section.
WTextBody textbody = document.getSections().get(0).getBody();
//Get the table.
IWTable table = textbody.getTables().get(0);
//Set the horizontal and vertical position for table.
table.getTableFormat().getPositioning().setHorizPosition(40);
table.getTableFormat().getPositioning().setVertPosition(100);
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

  
  
## How to set the text direction to a table in Word document?

The contents of the table cell can be in vertical or horizontal direction. Each cell content can have different text direction. The following code illustrates how to set the text direction for the text in the table.

{% tabs %}   

{% highlight JAVA %}
//Load the template document.
WordDocument document = new WordDocument("Template.docx");
//Get the text body of first section.
WTextBody textbody = document.getSections().get(0).getBody();
//Get the table
IWTable table = textbody.getTables().get(0);
//Iterate through table rows
for(Object row_tempObj : table.getRows())
{
	WTableRow row = (WTableRow)row_tempObj;
	for(Object cell_tempObj : row.getCells())
	{
		WTableCell cell = (WTableCell)cell_tempObj;
		//Set the text direction for the contents.
		cell.getCellFormat().setTextDirection(TextDirection.Vertical);
	}
}
//Save and close the document.
document.save("Sample.docx",FormatType.Docx);
document.close();
{% endhighlight %}

 {% endtabs %} 

 
 
## How to extract the images in the document?

The following code illustrates how to extract the images in the document.

{% tabs %} 

{% highlight JAVA %}


//Loads the template document

WordDocument document = new WordDocument("Template.docx");

WTextBody textbody = document.Sections[0].Body;

Image image;

int i = 1;

//Iterates through the paragraphs

foreach (WParagraph paragraph in textbody.Paragraphs)

{

//Iterates through the paragraph items 

foreach (ParagraphItem item in paragraph.ChildEntities)

{

//Gets the picture and saves it into specified location

switch (item.EntityType)

{

case EntityType.Picture:

WPicture picture = item as WPicture;

image = picture.Image;

image.Save(@"D:\Data\Image" + i + ".jpeg", ImageFormat.Jpeg);

i++;

break;

default:

break;

}

}

}

//Closes the document

document.Close();



{% endhighlight %}

{% highlight vb.net %}

'Loads the template document

Dim document As New WordDocument("Template.docx")

Dim textbody As WTextBody = document.Sections(0).Body

Dim image As Image

Dim i As Integer = 1

'Iterates through the paragraphs

For Each paragraph As WParagraph In textbody.Paragraphs

'Iterates through the paragraph items 

For Each item As ParagraphItem In paragraph.ChildEntities

'Gets the picture and saves it into specified location

Select Case item.EntityType

Case EntityType.Picture

Dim picture As WPicture = TryCast(item, WPicture)

image = picture.Image

image.Save("D:\Data\Image" & i & ".jpeg", ImageFormat.Jpeg)

i += 1

Exit Select

Case Else

Exit Select

End Select

Next

Next

'Close the document

document.Close()

{% endhighlight %} 

  {% endtabs %}  

The images in the document can be extracted into a specific location when exporting it to HTML file. The following code illustrates how to extract images.

{% tabs %}  

{% highlight JAVA %}

//Loads the template document

WordDocument document = new WordDocument("Template.docx");

//Sets the location to extract images

document.SaveOptions.HtmlExportImagesFolder = @"D:\Data\";

//Saves the document as html file

HTMLExport export = new HTMLExport();

export.SaveAsXhtml(document, "Template.html");

//Closes the document

document.Close();



{% endhighlight %}

{% endtabs %}  


## How to remove headers and footers from the document?

The following code illustrates how to remove the header contents from the document.

{% tabs %}  

{% highlight JAVA %}
//Load the template document.
WordDocument document = new WordDocument("Template.docx",FormatType.Docx);
//Iterate through the sections.
for(Object section_tempObj : document.getSections())
{
	WSection section = (WSection)section_tempObj;
	HeaderFooter header;
	//Get even footer of current section.
	header=section.getHeadersFooters().get(HeaderFooterType.EvenHeader);
	//Remove even footer.
	header.getChildEntities().clear();
	//Get odd footer of current section.
	header=section.getHeadersFooters().get(HeaderFooterType.OddHeader);
	//Remove odd footer.
	header.getChildEntities().clear();
	//Get first page footer.
	header=section.getHeadersFooters().get(HeaderFooterType.FirstPageHeader);
	//Remove first page footer.
	header.getChildEntities().clear();
}
//Save and close the document.
document.save("Sample.docx",FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

The following code illustrates how to remove the footer contents from the document.

{% tabs %}  

{% highlight JAVA %}
//Load the template document.
WordDocument document = new WordDocument("Template.docx");
//Iterate through the sections.
for(Object section_tempObj : document.getSections())
{
	WSection section = (WSection)section_tempObj;
	HeaderFooter footer;
	//Get even footer of current section.
	footer=section.getHeadersFooters().get(HeaderFooterType.EvenFooter);
	//Remove even footer.
	footer.getChildEntities().clear();
	//Get odd footer of current section.
	footer=section.getHeadersFooters().get(HeaderFooterType.OddFooter);
	//Remove odd footer.
	footer.getChildEntities().clear();
	//Get first page footer.
	footer=section.getHeadersFooters().get(HeaderFooterType.FirstPageFooter);
	//Remove first page footer.
	footer.getChildEntities().clear();
}
//Save and close the document.
document.save("Sample.docx",FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  
  
  
## Which units does Java Word library uses for measurement properties such as size, margins, etc, in a Word document?

Java Word library uses Points for measurement properties in a Word document.

## Migration from Microsoft Office Automation to Essential DocIO

### Bookmarks

Bookmarks identify the location of text in a Word document that you can name and identify for future reference.

Using Microsoft Office Automation

The following code example illustrates how to insert a bookmark for a range of text by using Office Automation.

{% tabs %}  

{% highlight JAVA %}

using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object nullobject = Missing.Value;

object newFilePath = "Sample.docx";

//Starts a Word application.

Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

//Creates a new Word document.

wordApp.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);

Microsoft.Office.Interop.Word.Document document = wordApp.ActiveDocument;

//Adds a paragraph to the document.

Microsoft.Office.Interop.Word.Paragraph oPara1;

oPara1 = document.Content.Paragraphs.Add(ref nullobject);

oPara1.Range.Text = "Bookmark with one word selected";

//Defines start and end positions of bookmark range.

object start = oPara1.Range.Text.IndexOf("word");

object end = oPara1.Range.Text.LastIndexOf(" ");

object rng = document.Range(ref start, ref end);

//Adds bookmark.

document.Bookmarks.Add("one_word", ref rng);

//Saves document and quits application.

document.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject);

//Closes document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

----------------

‘Initializes objects.

Dim nullobject As Object = Missing.Value

Dim newFilePath As Object = "Sample.docx"

‘Starts a Word application.

Dim wordApp As word.Application = New word.Application()

‘Creates a new Word document.

wordApp.Documents.Add(nullobject, nullobject, nullobject, nullobject)

Dim doc As word.Document = wordApp.ActiveDocument

‘Adds a paragraph to the document.

Dim oPara As word.Paragraph

oPara = doc.Content.Paragraphs.Add(nullobject)

oPara.Range.Text = "Bookmark with one word selected"

‘Defines the start and end positions of bookmark range.

Dim startobj As Object = oPara.Range.Text.IndexOf("word")

Dim endobj As Object = oPara.Range.Text.LastIndexOf(" ")

Dim rng As Object = doc.Range(startobj, endobj)

‘Adds bookmark.

doc.Bookmarks.Add("one_word", rng)

‘Saves document.

doc.SaveAs(newFilePath)

‘Closes document.

doc.Close(nullobject, nullobject, nullobject)

‘Quits application.

wordApp.Quit()

{% endhighlight %} 

 {% endtabs %}  

 
 
### Using DocIO

The following code example illustrates how to insert the bookmark by using DocIO. Here, the `appendBookmarkStart()` and `appendBookmarkEnd()` methods are used to add the bookmark.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document.
WordDocument doc = new WordDocument();
//Add new section.
IWSection section = doc.addSection();
//Add new paragraph.
IWParagraph paragraph = section.addParagraph();
paragraph.appendText("Simple Bookmark");
paragraph=section.addParagraph();
paragraph.appendText("Bookmark with one ");
//Insert bookmark.
paragraph.appendBookmarkStart("one_word");
paragraph.appendText("word");
paragraph.appendBookmarkEnd("one_word");
paragraph.appendText(" selected");
//Save the document.
doc.save("Sample.docx",FormatType.Docx);
//Close the document.
doc.close();
{% endhighlight %}

{% endtabs %}  


### Page Numbers

Page numbers can be added to the Word document in headers or footers.

Using Microsoft Office Automation

The following code example illustrates how page numbers can be inserted to the footer of the Word document by adding a page number field.

{% tabs %}   

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object filepath = "Sample.docx";

object nullobject = Missing.Value;

//Starts the Word application.

word.Application wordApp = new word.Application();

//Opens the Word document.

word.Document document = wordApp.Documents.Open(ref filepath, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject);

wordApp.Visible = false;

document.Activate();

//Seeks the page footer.

wordApp.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;

//Formats the footer.

wordApp.Selection.Paragraphs.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

wordApp.ActiveWindow.Selection.Font.Name = "Arial";

wordApp.ActiveWindow.Selection.Font.Size = 8;

//Adds page numbers in the footer.

Object CurrentPage = word.WdFieldType.wdFieldPage;

wordApp.ActiveWindow.Selection.Fields.Add(wordApp.Selection.Range, ref CurrentPage, ref nullobject, ref nullobject);

//Saves the document.

document.Save();

//Closes the document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

----------------

‘Initializes objects.

Dim nullobject As Object = Missing.Value

Dim filePath As Object =  "Sample.docx"

Dim falseobj As Object = False

‘Starts the application.

Dim wordApp As word.Application = New word.Application()

‘Adds a new Word document.

Dim document As word.Document = wordApp.Documents.Open(filePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, falseobj, nullobject, nullobject, nullobject, nullobject)

wordApp.Visible = False

document.Activate()

‘Seeks the page footer.

wordApp.ActiveWindow.ActivePane.View.SeekView = word.WdSeekView.wdSeekCurrentPageFooter

‘Formats the footer.

wordApp.Selection.Paragraphs.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter

wordApp.ActiveWindow.Selection.Font.Name = "Arial"

wordApp.ActiveWindow.Selection.Font.Size = 8

‘Adds page numbers in the footer.

Dim CurrentPage As Object = word.WdFieldType.wdFieldPage

wordApp.ActiveWindow.Selection.Fields.Add(wordApp.Selection.Range, CurrentPage, nullobject, nullobject)

‘Saves the document.

document.Save()

‘Closes the document.

document.Close(nullobject, nullobject, nullobject)

‘Quits application.

wordApp.Quit()

{% endhighlight %}

{% endtabs %} 

  
  
### Using DocIO

The following code example illustrates how page numbers are inserted to the footer of the Word document by using DocIO.

{% tabs %}   

{% highlight JAVA %}
//Open the Word document.
WordDocument doc = new WordDocument("Template.docx",FormatType.Docx);
//Iterate through sections.
for(Object sec_tempObj : doc.getSections())
{
	WSection sec = (WSection)sec_tempObj;
	IWParagraph para = sec.addParagraph();
	//Append page field to the paragraph.
	para.appendField("footer",FieldType.FieldPage);
	para.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
	sec.getPageSetup().setPageNumberStyle(PageNumberStyle.Arabic);
	//Add paragraph to footer.
	sec.getHeadersFooters().getFooter().getParagraphs().add(para);
}
//Save the document.
doc.save("Sample.docx",FormatType.Docx);
//Close the document.
doc.close();
{% endhighlight %}

{% endtabs %}   
  
### Document Watermarks

Watermarks are text or pictures that appear behind document text.

Using Microsoft Office Automation

The following code example illustrates how to insert a text watermark as a shape by using Office Automation.

{% tabs %}   

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object nullobject = Missing.Value;

object newFilePath = "Sample.docx";

//Starts the Word application.

word.Application wordApp = new word.Application();

//Creates a new Word document.

wordApp.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);

word.Document document = wordApp.ActiveDocument;

//Seeks the current page header.

wordApp.ActiveWindow.ActivePane.View.SeekView = word.WdSeekView.wdSeekCurrentPageHeader;

//Inserts watermark.

word.Shape watermark = wordApp.Selection.HeaderFooter.Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,

"Watermark", "Arial", (float)48, Microsoft.Office.Core.MsoTriState.msoTrue,

Microsoft.Office.Core.MsoTriState.msoFalse, 0, 0, ref nullobject);

//Sets watermark properties.

watermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

watermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

watermark.Fill.Solid();

watermark.Fill.ForeColor.RGB = (Int32)word.WdColor.wdColorGray30;

//Sets focus back to the document.

wordApp.ActiveWindow.ActivePane.View.SeekView = word.WdSeekView.wdSeekMainDocument;

//Saves the document.

document.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject);

//Closes the document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

----------------

Initializes objects.

Dim nullobject As Object = Missing.Value

Dim newFilePath As Object = "Sample.docx"

‘Starts the application.

Dim wordApp As word.Application = New word.Application()

‘Creates a new Word document.

wordApp.Documents.Add(nullobject, nullobject, nullobject, nullobject)

Dim doc As word.Document = wordApp.ActiveDocument

‘Seeks the current page header.

wordApp.ActiveWindow.ActivePane.View.SeekView = word.WdSeekView.wdSeekCurrentPageHeader

‘Adds text watermark to the document.

Dim watermark As word.Shape = wordApp.Selection.HeaderFooter.Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,"Watermark", "Arial", 48, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, 0, 0, nullobject)

‘Sets watermark properties.

watermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue

watermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse

watermark.Fill.Solid()

watermark.Fill.ForeColor.RGB = CType(word.WdColor.wdColorGray30, Integer)

‘Saves the document.

doc.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

‘Closes the document.

doc.Close(nullobject, nullobject, nullobject)

‘Quits application.

wordApp.Quit()

{% endhighlight %} 

  {% endtabs %} 

### Using DocIO

DocIO enables you to add a text watermark and a picture watermark to a Word document. The following code example shows how to insert the picture watermark to the Word document.

{% tabs %}  

{% highlight JAVA %}


//Creates a new Word document.

WordDocument doc = new WordDocument();

doc.EnsureMinimal();

//Adds picture watermark to the document.

PictureWatermark picWatermark = new PictureWatermark();

picWatermark.Scaling = 120f;

picWatermark.Washout = true;

doc.Watermark = picWatermark;

picWatermark.Picture = Image.FromFile(ImagesPath + "Water lilies.jpg");

//Saves the document.

doc.Save("Sample.docx", FormatType.Docx);

//Closes the document.

doc.Close();



{% endhighlight %}

{% highlight vb.net %}

‘Creates a new Word document.

Dim doc As WordDocument = New WordDocument()

doc.EnsureMinimal()

‘Adds picture watermark to the document.

Dim picWatermark As PictureWatermark = New PictureWatermark()

picWatermark.Scaling = 120f

picWatermark.Washout = True

doc.Watermark = picWatermark

picWatermark.Picture = Image.FromFile(ImagesPath and "Water lilies.jpg")

‘Saves the document.

doc.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

doc.Close()

{% endhighlight %} 

 {% endtabs %}  

N>  For more information on adding watermarks to a Word document using DocIO, refer to the online documentation link:
[Applying Watermark](/File-Formats/DocIO/Applying-Watermark)

### Headers and Footers

The headers and footers can be inserted with text, graphics, and any other information that is contained in the document. 

Using Microsoft Office Automation 

The following code example illustrates how to add headers and footers to a Word document. In this example, page numbers are inserted to the header and a text is inserted to the footer.

{% tabs %}  

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object nullobject = Missing.Value;

object filePath = "Template.docx";

object newFilePath = "Sample.docx";

//Starts the Word application.

word.Application wordApp = new word.Application();

//Opens the Word document.

word.Document document = wordApp.Documents.Open(ref filePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, ref nullobject);

wordApp.Visible = false;

//Adds header and footer to each section in the document.

foreach (word.Section section in document.Sections)

{

object fieldEmpty = word.WdFieldType.wdFieldPage;

object autoText = "AUTOTEXT  \"Page X of Y\" ";

object preserveFormatting = true;

//Footer.

section.Footers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "Internal";

section.Footers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = 

word.WdParagraphAlignment.wdAlignParagraphLeft;

//Header.       

section.Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Add(section.Headers[

word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range,reffieldEmpty, ref autoText, ref preserveFormatting);

section.Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = 

word.WdParagraphAlignment.wdAlignParagraphRight;

}

//Saves the document.

document.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject);

//Closes the document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

----------------

'Initializes objects.

Dim nullobject As Object = System.Reflection.Missing.Value

Dim filePath As Object = "Template.docx"

Dim newFilePath As Object = "Sample.docx"

'Starts the application.

Dim wordApp As word.Application = New word.Application()

'Opens the document.

Dim document As word.Document = wordApp.Documents.Open(filePath, nullobject, nullobject, nullobject, nullobject, nullobject,

nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

wordApp.Visible = False

'Adds header and footer to each section in the document.

For Each section As word.Section In document.Sections

Dim fieldEmpty As Object = word.WdFieldType.wdFieldPage

'Footer.    

section.Footers(word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text = "Internal"

section.Footers(word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphLeft

'Header.       

section.Headers(word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Fields.Add(section.Headers(word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range, fieldEmpty, nullobject, nullobject)

section.Headers(word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphRight

Next

'Saves the document.

document.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

'Closes the document.

document.Close(nullobject, nullobject, nullobject)

'Quits the application.

wordApp.Quit()

{% endhighlight %}

  {% endtabs %}  

### Using DocIO

You can set the header and footer by using the HeadersFooters property in the Word document section. To access a particular header/footer, you can use the following properties of `WHeadersFooters` class:

* FirstPageHeader
* FirstPageFooter
* OddHeader
* OddFooter
* EvenHeader
* EvenFooter


{% tabs %}   

{% highlight JAVA %}
//Open a Word document.
WordDocument doc = new WordDocument("Template.docx");
//Add header and footer to each section in the document.
for(Object sec_tempObj : doc.getSections())
{
	//Header.
	WSection sec = (WSection)sec_tempObj;
	WParagraph para = new WParagraph(doc);
	para.appendField("page",FieldType.FieldPage);
	para.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
	sec.getHeadersFooters().getHeader().getParagraphs().add(para);
	//Footer.
	WParagraph para1 = new WParagraph(doc);
	para1.appendText("Internal");
	para1.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
	sec.getHeadersFooters().getFooter().getParagraphs().add(para1);
}
//Save the document.
doc.save("Sample.docx",FormatType.Docx);
//Close the document.
doc.close();
{% endhighlight %}

{% endtabs %} 


### Character Formatting

Character formatting defines the appearance of the text in a Word document. This section illustrates how to apply character level formatting to the Word document. 

Using Microsoft Office Automation

The following code example illustrates how to apply the character formatting to the Word document by using the Range properties.

{% tabs %} 

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word

----------------

//Initializes objects.

object nullobject = System.Reflection.Missing.Value;

object newFilePath = "Sample.docx";

object falseObj = false;

//Starts the Word application.

word.Application wordApp = new word.Application();

//Creates a new Word document.

wordApp.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);

word.Document doc = wordApp.ActiveDocument;

//Defines the range for formatting.

object start = 0;

object end = 0;

word.Range rng = doc.Range(ref start, ref end);

rng.Text = "New Text";

rng.Font.Name = "Arial";

rng.Font.Size = 14;

//Saves the document.

doc.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject);

//Closes the document.

doc.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

----------------

‘Initializes objects.

Dim nullobject As Object = System.Reflection.Missing.Value

Dim newFilePath As Object = "Sample.docx"

Dim falseObj As Object = False

‘Starts the Word application.

Dim wordApp As word.Application = New word.Application()

‘Creates a new Word document.

wordApp.Documents.Add(nullobject, nullobject, nullobject, nullobject)

Dim doc As word.Document = wordApp.ActiveDocument

‘Defines the range for formatting.

Dim start As Object = 0

Dim endobj As Object = 0

Dim rng As word.Range = doc.Range(start, endobj)

rng.Text = "New Text"

rng.Font.Name = "Arial"

rng.Font.Size = 14

‘Saves the document.

doc.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

‘Closes the document.

doc.Close(nullobject, nullobject, nullobject)

‘Quits the application.

wordApp.Quit()

{% endhighlight %} 

{% endtabs %}  

  
  
### Tables

Tables are used to organize information and to display the information in rows and columns. You can also add images or even other tables to the table.

Using Microsoft Office Automation

The following code example illustrates how to insert a table to a Word document, where the table contains three rows and two columns.

{% tabs %}  

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes the objects.

object nullobject = System.Reflection.Missing.Value;

object newFilePath = "Sample.docx";

//Starts the Word application.

word.Application wordApp = new word.Application();

//Creates a new document.

wordApp.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);

word.Document document = wordApp.ActiveDocument;

//Inserts the table.

object start = 0;

object end = 0;

word.Range tableLocation = document.Range(ref start, ref end);

document.Tables.Add(tableLocation, 3, 2, ref nullobject, ref nullobject);

//Saves the document.

document.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject);

//Closes the document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

---------

'Initializes the objects.

Dim nullobject As Object = System.Reflection.Missing.Value

Dim newFilePath As Object = "Sample.docx"

'Starts the Word application.

Dim wordApp As New word.Application()

'Creates a new document.

wordApp.Documents.Add(nullobject, nullobject, nullobject, nullobject)

Dim document As word.Document = wordApp.ActiveDocument

'Inserts the table.

Dim start As Object = 0

Dim [end] As Object = 0

Dim tableLocation As word.Range = document.Range(start, [end])

document.Tables.Add(tableLocation, 3, 2, nullobject, nullobject)

'Saves the document.

document.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, _

nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, _

nullobject, nullobject, nullobject, nullobject)

'Closes the document.

document.Close(nullobject, nullobject, nullobject)

'Quits the application.

wordApp.Quit(nullobject, nullobject, nullobject)

{% endhighlight %} 

{% endtabs %}  

### Using DocIO

The following code example shows how to insert an empty table to a Word document. The `resetCells()` method is used to specify the number of rows and columns in a table.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
//Add a table to the document.
IWTable table = section.addTable();
table.resetCells(3, 2);
//Save the document.
document.save("Sample.docx",FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %} 

{% endtabs %}  
   
N>  For more information on creating tables using DocIO, refer to online documentation link:
[Working with Tables](https://help.syncfusion.com/java-file-formats/word-library/working-with-tables)


### Comments 

Comments are used to include additional information to a paragraph or text in a Word document. Comments can be added or modified whenever needed and deleted when the comment has served its purpose. 

Adding Comments using Microsoft Office Automation

The following code example illustrates how to add comments to a Word document. You need to define the range of text where the comment is to be added.

{% tabs %}  

{% highlight JAVA %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object nullobject = System.Reflection.Missing.Value;

object newFilePath = "Sample.docx";

//Starts the Word application.

word.Application wordApp = new word.Application();

//Creates a new document.

wordApp.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);

word.Document doc = wordApp.ActiveDocument;

//Inserts text to the Word document.

object start = 0;

object end = 0;

word.Range rng = doc.Range(ref start, ref end);

rng.Text = "New Text";

//Adds comment to the inserted text.

object text = "Comment goes here";

doc.Comments.Add(rng, ref text);

//Saves the document.

doc.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject);

//Closes the document.

doc.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}


Imports word = Microsoft.Office.Interop.Word

---------

‘Initializes objects.

Dim nullobject As Object = System.Reflection.Missing.Value

Dim newFilePath As Object = "Sample.docx"

‘Starts the Word application.

Dim wordApp As word.Application = New word.Application()

‘Creates a new document.

wordApp.Documents.Add(nullobject, nullobject, nullobject, nullobject)

Dim doc As word.Document = wordApp.ActiveDocument

‘Inserts text to the Word document.

Dim startobj As Object = 0

Dim endobj As Object = 0

Dim rng As word.Range = doc.Range(startobj, endobj)

rng.Text = "New Text"

‘Adds comment to the inserted text.

Dim text As Object = "Comment goes here"

doc.Comments.Add(rng, text)

‘Saves the document and quits application.

doc.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

‘Closes the document.

doc.Close(nullobject, nullobject, nullobject)

‘Quits the application.

wordApp.Quit()

{% endhighlight %} 

 {% endtabs %}  

#### Adding Comments Using DocIO

You can insert comments to a paragraph or text in a Word document by using DocIO. The following code example shows how to insert comments to a Word document.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document.
WordDocument doc = new WordDocument();
IWSection section = doc.addSection();
//Add a paragraph to the document.
IWParagraph para = section.addParagraph();
para.appendText("New Text");
//Add comment to the paragraph.
para.appendComment("Comment goes here");
//Save the document.
doc.save("Sample.docx", FormatType.Docx);
{% endhighlight %}

{% endtabs %} 

N>  For more information on working with the comments using Java Word library, you can refer to the online documentation link:
[Working with Comments](https://help.syncfusion.com/java-file-formats/word-library/working-with-comments) 


## How to check whether a Word document contains tracked changes or not? 

You can check whether a Word document contains tracked changes by using `HasChanges` property in Essential DocIO.

The following code example shows how to check whether a Word document contains tracked changes.

{% tabs %}   

{% highlight JAVA %}
//Open an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Get a flag which denotes whether the Word document has track changes.
boolean hasChanges = document.getHasChanges();
//When the document has track changes, accepts all changes.
if (hasChanges)
	document.getRevisions().acceptAll();
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

## How to accept or reject track changes of specific type in the Word document?

You can **accept or reject track changes by revision type** in the tracked changes Word document. 

For example, if you like to accept or reject changes of specific revision type (insertions, deletions, formatting, move to, or move from), you can iterate into the revisions in Word document and then accept or reject the particular revision type using Essential DocIO.

The following code example shows how to accept or reject track changes of specific type in the Word document .

{% tabs %}   

{% highlight JAVA %}
//Open an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Iterate into all the revisions in Word document
for (int i = document.getRevisions().getCount() - 1; i >= 0; i--)
{
	// Get the type of the track changes revision.
	RevisionType revisionType = document.getRevisions().get(i).getRevisionType();
	//Accept only insertion and Move from revisions changes.
	if (revisionType == RevisionType.Insertions || revisionType == RevisionType.MoveFrom)
		document.getRevisions().get(i).accept();
	//Reset to last item when accept the moving related revisions.
	if (i > document.getRevisions().getCount() - 1)
		i = document.getRevisions().getCount();
}
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

## How to enable track changes for Word document?

TrackChanges is used to keep track of the changes made to a Word document. This can be enabled by using the TrackChanges property of the Word document.

The following code example shows how to enable track changes of the document.

{% tabs %}   

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Append text to the paragraph.
IWTextRange text = paragraph.appendText("This sample illustrates how to track the changes made to the word document. ");
//Set font name and size for text.
text.getCharacterFormat().setFontName("Times New Roman");
text.getCharacterFormat().setFontSize(14);
text = paragraph.appendText("This track changes is useful in shared environment.");
text.getCharacterFormat().setFontSize(12);
//Turn on the track changes option.
document.setTrackChanges(true);
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

