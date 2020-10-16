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

{% highlight c# %}


//Loads the template document

WordDocument document = new WordDocument("Template.docx");

//Gets the text body of first section

WTextBody textbody = document.Sections[0].Body;

//Gets the table

IWTable table = textbody.Tables[0];

//Iterates through table rows

foreach (WTableRow row in table.Rows)

{

foreach (WTableCell cell in row.Cells)

{

//Sets the text direction for the contents

cell.CellFormat.TextDirection = Syncfusion.DocIO.DLS.TextDirection.Vertical;

}

}

//Saves and closes the document

document.Save("Sample.docx", FormatType.Docx);

document.Close();



{% endhighlight %}

{% highlight vb.net %}

'Loads the template document

Dim document As New WordDocument("Template.docx")

'Gets the text body of first section

Dim textbody As WTextBody = document.Sections(0).Body

'Gets the table

Dim table As IWTable = textbody.Tables(0)

'Iterates through table rows

For Each row As WTableRow In table.Rows

For Each cell As WTableCell In row.Cells

'Sets the text direction for the contents

cell.CellFormat.TextDirection = Syncfusion.DocIO.DLS.TextDirection.Vertical

Next

Next

'Saves and closes the document

document.Save("Sample.docx", FormatType.Docx)

document.Close()

{% endhighlight %} 

 {% endtabs %} 

 
 
## How to extract the images in the document?

The following code illustrates how to extract the images in the document.

{% tabs %} 

{% highlight c# %}


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

{% highlight c# %}

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

{% highlight vb.net %}

'Loads the template document

Dim document As New WordDocument("Template.docx")

'Sets the location to extract images

document.SaveOptions.HtmlExportImagesFolder = "D:\Data\"

'Saves the document as html file

Dim export As New HTMLExport()

export.SaveAsXhtml(document, "Template.html")

'Closes the document

document.Close()



{% endhighlight %}

{% endtabs %}  


## How to remove headers and footers from the document?

The following code illustrates how to remove the header contents from the document.

{% tabs %}  

{% highlight c# %}


//Loads the template document

WordDocument document = new WordDocument("Template.docx", FormatType.Docx);

//Iterates through the sections

foreach (WSection section in document.Sections)

{

HeaderFooter header;

//Gets even footer of current section

header = section.HeadersFooters[HeaderFooterType.EvenHeader];

//Removes even footer

header.ChildEntities.Clear();

//Gets odd footer of current section

header = section.HeadersFooters[HeaderFooterType.OddHeader];

//Removes odd footer

header.ChildEntities.Clear();

//Gets first page footer

header = section.HeadersFooters[HeaderFooterType.FirstPageHeader];

//Removes first page footer

header.ChildEntities.Clear();

}

//Saves and closes the document

document.Save("Sample.docx", FormatType.Docx);

document.Close();



{% endhighlight %}

{% highlight vb.net %}

'Loads the template document

Dim document As New WordDocument("Template.docx", FormatType.Docx)

'Iterates through the sections

For Each section As WSection In document.Sections

Dim header As HeaderFooter

'Gets even footer of current section

header = section.HeadersFooters(HeaderFooterType.EvenHeader)

'Removes even footer

header.ChildEntities.Clear()

'Gets odd footer of current section

header = section.HeadersFooters(HeaderFooterType.OddHeader)

'Removes odd footer

header.ChildEntities.Clear()

'Gets first page footer

header = section.HeadersFooters(HeaderFooterType.FirstPageHeader)

'Removes first page footer

header.ChildEntities.Clear()

Next

'Saves and closes the document

document.Save("Sample.docx", FormatType.Docx)

document.Close()

{% endhighlight %} 

 {% endtabs %}  

The following code illustrates how to remove the footer contents from the document.

{% tabs %}  

{% highlight c# %}


//Loads the template document

WordDocument document = new WordDocument("Template.docx");            

//Iterates through the sections

foreach (WSection section in document.Sections)

{

HeaderFooter footer;

//Gets even footer of current section

footer = section.HeadersFooters[HeaderFooterType.EvenFooter];

//Removes even footer

footer.ChildEntities.Clear();

//Gets odd footer of current section

footer = section.HeadersFooters[HeaderFooterType.OddFooter];

//Removes odd footer

footer.ChildEntities.Clear();

//Gets first page footer

footer = section.HeadersFooters[HeaderFooterType.FirstPageFooter];

//Removes first page footer

footer.ChildEntities.Clear();

}

//Saves and closes the document

document.Save("Sample.docx", FormatType.Docx);

document.Close();



{% endhighlight %}

{% highlight vb.net %}

'Loads the template document

Dim document As New WordDocument("Template.docx")

'Iterates through the sections

For Each section As WSection In document.Sections

Dim footer As HeaderFooter

'Gets even footer of current section

footer = section.HeadersFooters(HeaderFooterType.EvenFooter)

'Removes even footer

footer.ChildEntities.Clear()

'Gets odd footer of current section

footer = section.HeadersFooters(HeaderFooterType.OddFooter)

'Removes odd footer

footer.ChildEntities.Clear()

'Gets first page footer

footer = section.HeadersFooters(HeaderFooterType.FirstPageFooter)

'Removes first page footer

footer.ChildEntities.Clear()

Next

'Saves and closes the document

document.Save("Sample.docx", FormatType.Docx)

document.Close()

{% endhighlight %}

  {% endtabs %}  

  
  
## Which units does Essential DocIO uses for measurement properties such as size, margins, etc, in a Word document?

Essential DocIO library uses Points for measurement properties in a Word document.

##  Could not find Syncfusion.OfficeChartToImageConverter assembly in .NET 3.5 Framework, does it mean there is no support for chart conversion in this Framework? 

Yes, OfficeChartToImageConverter assembly is not supported in .NET 3.5 Framework and it is available in .NET 4.0 Framework.

## Can the chart data be refreshed?

Yes, Essential DocIO supports refreshing the chart data. For more details, refer [Working with charts](/File-Formats/DocIO/Working-with-Charts)

## Is it possible to convert 3D charts to PDF or image?

Current version of the DocIO library does not provide support for converting 3D charts to PDF or image format.

## Is it possible to specify PDF conformance level in Word to PDF conversion?

Yes, you can specify the PDF conformance level in Word to PDF conversion. For more details, refer [PDF Conformance](/file-formats/pdf/working-with-pdf-conformance)

## Migration from Microsoft Office Automation to Essential DocIO

### Mail merge

The Mail merge feature can be used to generate reports and letters in Microsoft Word. The following code examples show how to generate an employee report from an MDB data source by using Office Automation and DocIO.

Using Microsoft Office Automation

Office Automation performs the Mail merge by executing a SQL query on the Word document. The output of the Mail merge can be sent to a new Word document. Alternatively, it can be sent to a printer, a fax machine, or forwarded to an e-mail address.

{% tabs %}  

{% highlight c# %}

using word = Microsoft.Office.Interop.Word;

------------

//Initializes objects.

object nullobject = Missing.Value;

object filepath = "Sample.docx";

object sqlStmt = "SELECT * FROM [Employees]";

string sDBPath = "Northwind.mdb";

//Starts the Word application.

word.Application wordApp = new word.Application();

//Opens the Word document.

word.Document document = wordApp.Documents.Open(ref filepath, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject);

wordApp.Visible = false;

//Performs Mail Merge.     

document.MailMerge.OpenDataSource(sDBPath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref sqlStmt, ref nullobject,

ref nullobject, ref nullobject);

document.MailMerge.Execute(ref nullobject);

//Sends output of Mail Merge to a new document.

document.MailMerge.Destination = word.WdMailMergeDestination.wdSendToNewDocument;

//Closes the document.

document.Close(ref nullobject, ref nullobject, ref nullobject);

//Quits the application.

wordApp.Quit(ref nullobject, ref nullobject, ref nullobject);



{% endhighlight %}

{% highlight vb.net %}

Imports word = Microsoft.Office.Interop.Word

-------------

'Initializes objects.

Dim nullobject As Object = Missing.Value

Dim filepath As Object = "Sample.docx"

Dim sqlStmt As Object = "SELECT * FROM [Employees]"

Dim sDBPath As String = "Northwind.mdb"

'Starts the Word application.

Dim wordApp As New word.Application()

'Opens the Word document.

Dim document As word.Document = wordApp.Documents.Open(filepath, nullobject, nullobject, nullobject, nullobject, nullobject, _

nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, _

nullobject, nullobject, nullobject, nullobject)

wordApp.Visible = False

'Performs Mail Merge.     

document.MailMerge.OpenDataSource(sDBPath, nullobject, nullobject, nullobject, nullobject, nullobject, _

nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, _

sqlStmt, nullobject, nullobject, nullobject)

document.MailMerge.Execute(nullobject)

'Sends output of Mail Merge to a new document.

document.MailMerge.Destination = word.WdMailMergeDestination.wdSendToNewDocument

'Closes the document.

document.Close(nullobject, nullobject, nullobject)

'Quits the application.

wordApp.Quit(nullobject, nullobject, nullobject)

{% endhighlight %}

 {% endtabs %}  
 
 

### Using DocIO

DocIO performs Mail merge by using the following methods:

* Execute
* ExecuteGroup
* ExecuteNestedGroup

The following code example performs Mail merge by using the `Execute` method.

{% tabs %}    

{% highlight c# %}


string dataBase = "Northwind.mdb";

//Opens existing template.

WordDocument doc = new WordDocument("Template.docx", FormatType.Docx);

//Gets Data from the Database.

OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataBase);

conn.Open();

//Populates the data table.

DataTable table = new DataTable();

OleDbDataAdapter adapter = new OleDbDataAdapter("select * from employees", conn);

adapter.Fill(table);

adapter.Dispose();

//Performs Mail Merge.

doc.MailMerge.Execute(table);

//Saves the document.

doc.Save("Sample.docx", FormatType.Docx);

//Closes the document.

doc.Close();



{% endhighlight %}

{% highlight vb.net %}

Dim dataBase As String = "Northwind.mdb" 

‘Opens the Word document.

Dim doc As WordDocument = New WordDocument("Template.docx")

‘Creates database connection.

Dim conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataBase)

conn.Open()

‘Populates data table.

Dim table As DataTable = New DataTable()

Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("select * from employees", conn)

adapter.Fill(table)

adapter.Dispose()

‘Performs Mail Merge.

doc.MailMerge.Execute(table)

‘Saves the document.

doc.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

doc.Close()

{% endhighlight %}

  {% endtabs %}

N> 
For more information on Mail merge using DocIO, you can refer to online documentation link:
[MailMerge](/File-Formats/DocIO/Working-with-MailMerge)

### Find and Replace

This section illustrates how to perform a simple find and replace operation in a Word document by using Microsoft Office Automation and DocIO.

Using Microsoft Office Automation

The following code example illustrates how to search for a word in a Word document, replace it with another word and save the document under a new name.

{% tabs %}  

{% highlight c# %}


using word = Microsoft.Office.Interop.Word;

---------

//Initializes objects.

object nullobject = Missing.Value;

object filepath = "Template.docx";

object newFilePath = "Sample.docx";

object item = word.WdGoToItem.wdGoToPage;

object whichItem = word.WdGoToDirection.wdGoToFirst;

object replaceAll = word.WdReplace.wdReplaceAll;

object forward = true;

object matchAllWord = true;

object matchCase = false;

object originalText = "Hello";

object replaceText = "World";

object save = true;

//Starts the Word application.

word.Application wordApp = new word.Application();

//Opens the Word document.

word.Document document = wordApp.Documents.Open(ref filepath, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject,ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject);

wordApp.Visible = false;

//Searches and replaces text.

document.GoTo(ref item, ref whichItem, ref nullobject, ref nullobject);

foreach (word.Range rng in document.StoryRanges)

{

rng.Find.Execute(ref originalText, ref matchCase, ref matchAllWord, ref nullobject, ref nullobject,

ref nullobject, ref forward,ref nullobject, ref nullobject, ref replaceText, ref replaceAll,

ref nullobject, ref nullobject, ref nullobject, ref nullobject);

}

//Saves the document.

document.SaveAs(ref newFilePath, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, 

ref nullobject, ref nullobject, refnullobject, ref nullobject,

ref nullobject);

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

Dim filePath As Object = "Template.docx"

Dim newFilePath As Object = "Sample.docx"

Dim item As Object = word.WdGoToItem.wdGoToPage

Dim whichItem As Object = word.WdGoToDirection.wdGoToFirst

Dim replaceAll As Object = word.WdReplace.wdReplaceAll

Dim forward As Object = True

Dim matchAllWord As Object = True

Dim matchCase As Object = False

Dim originalText As Object = "Hello"

Dim replaceText As Object = "World"

Dim save As Object = True

Dim falseObj As Object = False

‘Starts the Word application.

Dim wordApp As word.Application = New word.Application()

‘Opens the Word document.

Dim doc As word.Document = wordApp.Documents.Open(filePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, falseobj, nullobject, nullobject, nullobject, nullobject)

wordApp.Visible = False

‘Searches and replaces text.

doc.GoTo(item, whichItem, nullobject, nullobject)

For Each rng As word.Range In doc.StoryRanges

rng.Find.Execute(originalText, matchCase, matchAllWord, nullobject, nullobject, nullobject, forward, nullobject, nullobject, replaceText, replaceAll, nullobject, nullobject, nullobject, nullobject)

Next

‘Saves the document.

doc.SaveAs(newFilePath, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject)

‘Closes the document.

doc.Close(nullobject, nullobject, nullobject)

‘Quits the application.

wordApp.Quit(nullobject, nullobject, nullobject)

{% endhighlight %} 

 {% endtabs %}  

 
 
### Using DocIO

The following code example illustrates how to perform a simple find and replace operation by using DocIO.

{% tabs %}  

{% highlight c# %}


//Opens the Word document.

WordDocument document = new WordDocument("Template.docx",FormatType.Docx);

//Defines replacement text.

string replaceText = "World";

//Performs replace.

document.Replace(new Regex("Hello"), replaceText);

//Saves the document.

document.Save("Sample.docx", FormatType.Docx);

//Closes the document.

document.Close();



{% endhighlight %}

{% highlight vb.net %}

‘Opens the Word document.

Dim document As WordDocument = New WordDocument("Template.docx")

‘Defines text to be replaced.

Dim replaceText As String = "World"

‘Performs replace.

document.Replace(New Regex("Hello"), replaceText)

‘Saves the document.

document.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

document.Close()

{% endhighlight %} 

 {% endtabs %}  


N>  For more information on performing the find and replace operation using DocIO, you can refer to online documentation link:
[Find and Replace](/File-Formats/DocIO/Working-with-Find-and-Replace)



### Bookmarks

Bookmarks identify the location of text in a Word document that you can name and identify for future reference.

Using Microsoft Office Automation

The following code example illustrates how to insert a bookmark for a range of text by using Office Automation.

{% tabs %}  

{% highlight c# %}

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

The following code example illustrates how to insert the bookmark by using DocIO. Here, the `AppendBookmarkStart()` and `AppendBookmarkEnd()` methods are used to add the bookmark.

{% tabs %}  

{% highlight c# %}


//Creates a new Word document.

WordDocument doc = new WordDocument();

//Adds new section

IWSection section = doc.AddSection();

//Adds new paragraph

IWParagraph paragraph = section.AddParagraph();

paragraph.AppendText("Simple Bookmark");

paragraph = section.AddParagraph();

paragraph.AppendText("Bookmark with one ");

//Inserts bookmark.

paragraph.AppendBookmarkStart("one_word");

paragraph.AppendText("word");

paragraph.AppendBookmarkEnd("one_word");

paragraph.AppendText(" selected");

//Saves the document.

doc.Save("Sample.docx", FormatType.Docx);

//Closes the document.

doc.Close();



{% endhighlight %}

{% highlight vb.net %}

‘Creates a new Word document.

Dim doc As WordDocument = New WordDocument()

‘Adds new section

Dim section As IWSection = doc.AddSection()

‘Adds new paragraph

Dim paragraph As IWParagraph = section.AddParagraph()

paragraph.AppendText("Simple Bookmark")

paragraph = section.AddParagraph()

paragraph.AppendText("Bookmark with one ")

‘Inserts bookmark.

paragraph.AppendBookmarkStart("one_word")

paragraph.AppendText("word")

paragraph.AppendBookmarkEnd("one_word")

paragraph.AppendText(" selected")

‘Saves the document.

doc.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

doc.Close()

{% endhighlight %}

  {% endtabs %}  



### Page Numbers

Page numbers can be added to the Word document in headers or footers.

Using Microsoft Office Automation

The following code example illustrates how page numbers can be inserted to the footer of the Word document by adding a page number field.

{% tabs %}   

{% highlight c# %}


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

{% highlight c# %}


//Opens the Word document.

WordDocument doc = new WordDocument("Template.docx", FormatType.Docx);

//Iterates through sections

foreach (WSection sec in doc.Sections)

{

IWParagraph para = sec.AddParagraph();

//Appends page field to the paragraph

para.AppendField("footer", FieldType.FieldPage);

para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

sec.PageSetup.PageNumberStyle = PageNumberStyle.Arabic;

//Adds paragraph to footer

sec.HeadersFooters.Footer.Paragraphs.Add(para);

}

//Saves the document.

doc.Save("Sample.docx",FormatType.Docx);

//Closes the document.

doc.Close();



{% endhighlight %}

{% highlight vb.net %}

‘Opens the Word document.

Dim doc As WordDocument = New WordDocument("Template.docx", FormatType.Docx)

‘Iterates through sections

For Each sec As WSection In doc.Sections

Dim para As IWParagraph = sec.AddParagraph()

‘Appends page field to the paragraph

para.AppendField("footer", FieldType.FieldPage)

para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center

sec.PageSetup.PageNumberStyle = PageNumberStyle.Arabic

‘Adds paragraph to footer

sec.HeadersFooters.Footer.Paragraphs.Add(para)

Next

‘Saves the document.

doc.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

doc.Close()

{% endhighlight %}

  {% endtabs %} 

  
  
### Document Watermarks

Watermarks are text or pictures that appear behind document text.

Using Microsoft Office Automation

The following code example illustrates how to insert a text watermark as a shape by using Office Automation.

{% tabs %}   

{% highlight c# %}


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

{% highlight c# %}


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

{% highlight c# %}


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

{% highlight c# %}


//Opens a Word document.

WordDocument doc = new WordDocument("Template.docx");

//Adds header and footer to each section in the document.

foreach (WSection sec in doc.Sections)

{

//Header.

WParagraph para = new WParagraph(doc);

para.AppendField("page", FieldType.FieldPage);

para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

sec.HeadersFooters.Header.Paragraphs.Add(para);

//Footer.

WParagraph para1 = new WParagraph(doc);

para1.AppendText("Internal");

para1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;

sec.HeadersFooters.Footer.Paragraphs.Add(para1);

}

//Saves the document.

doc.Save("Sample.docx", FormatType.Docx);

//Closes the document.

doc.Close();



{% endhighlight %}

{% highlight vb.net %}

‘Opens the Word document.

Dim doc As WordDocument = New WordDocument("Template.docx")

‘Adds header and footer to each section in the document.

For Each sec As WSection In doc.Sections

‘Header.

Dim para As WParagraph = New WParagraph(doc)

para.AppendField("page", FieldType.FieldPage)

para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right

sec.HeadersFooters.Header.Paragraphs.Add(para)

‘Footer.

Dim para1 As WParagraph = New WParagraph(doc)

para1.AppendText("Internal")

para1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left

sec.HeadersFooters.Footer.Paragraphs.Add(para1)

Next

‘Saves the document.

doc.Save("Sample.docx", FormatType.Docx)

‘Closes the document.

doc.Close()

{% endhighlight %} 

 {% endtabs %} 


### Character Formatting

Character formatting defines the appearance of the text in a Word document. This section illustrates how to apply character level formatting to the Word document. 

Using Microsoft Office Automation

The following code example illustrates how to apply the character formatting to the Word document by using the Range properties.

{% tabs %} 

{% highlight c# %}


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

{% highlight c# %}


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

The following code example shows how to insert an empty table to a Word document. The `ResetCells()` method is used to specify the number of rows and columns in a table.

{% tabs %} 

{% highlight c# %}


//Creates a new Word document.

WordDocument document = new WordDocument();

IWSection section = document.AddSection();

//Adds a table to the document.

IWTable table = section.AddTable();

table.ResetCells(3, 2);

//Saves the document.

document.Save("Sample.docx",FormatType.Docx);

//Closes the document.

document.Close();   

{% endhighlight %}

{% highlight vb.net %}

'Creates a new Word document.

Dim document As New WordDocument()

Dim section As IWSection = document.AddSection()

'Adds a table to the document.

Dim table As IWTable = section.AddTable()

table.ResetCells(3, 2)

'Saves the document.

document.Save("Sample.docx",FormatType.Docx);

'Closes the document.

document.Close()

{% endhighlight %} 

   {% endtabs %}  

   
N>  For more information on creating tables using DocIO, refer to online documentation link:
[Working with Tables](/File-Formats/DocIO/Working-with-Tables)


### Comments 

Comments are used to include additional information to a paragraph or text in a Word document. Comments can be added or modified whenever needed and deleted when the comment has served its purpose. 

Adding Comments using Microsoft Office Automation

The following code example illustrates how to add comments to a Word document. You need to define the range of text where the comment is to be added.

{% tabs %}  

{% highlight c# %}


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

{% highlight c# %}

//Creates a new Word document.

WordDocument doc = new WordDocument();

IWSection section = doc.AddSection();

//Adds a paragraph to the document.

IWParagraph para = section.AddParagraph();

para.AppendText("New Text");

//Adds comment to the paragraph.

para.AppendComment("Comment goes here");

//Saves the document.

doc.Save("Sample.doc", FormatType.Doc);


{% endhighlight %}

 {% endtabs %} 

N>  For more information on working with the comments using DocIO, you can refer to the online documentation link:
[Working with Comments](/File-Formats/DocIO/Working-with-Comments) 


## How to check whether a Word document contains tracked changes or not? 

You can check whether a Word document contains tracked changes by using `HasChanges` property in Essential DocIO.

The following code example shows how to check whether a Word document contains tracked changes.

{% tabs %}   

{% highlight c# %}
//Opens an existing Word document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Gets a flag which denotes whether the Word document has track changes
bool hasChanges = document.HasChanges;
//When the document has track changes, accepts all changes
if (hasChanges)
	document.Revisions.AcceptAll();
//Saves and closes the document
document.Save("Sample.docx", FormatType.Docx);
document.Close();
{% endhighlight %} 

{% endtabs %}

## How to accept or reject track changes of specific type in the Word document?

You can **accept or reject track changes by revision type** in the tracked changes Word document. 

For example, if you like to accept or reject changes of specific revision type (insertions, deletions, formatting, move to, or move from), you can iterate into the revisions in Word document and then accept or reject the particular revision type using Essential DocIO.

The following code example shows how to accept or reject track changes of specific type in the Word document .

{% tabs %}   

{% highlight c# %}
//Opens an existing Word document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Iterates into all the revisions in Word document
for (int i = document.Revisions.Count - 1; i >= 0; i--)
{
	// Gets the type of the track changes revision
	RevisionType revisionType = document.Revisions[i].RevisionType;
	//Accepts only insertion and Move from revisions changes
	if (revisionType == RevisionType.Insertions || revisionType == RevisionType.MoveFrom)
		document.Revisions[i].Accept();
	//Resets to last item when accept the moving related revisions.
	if (i > document.Revisions.Count - 1)
		i = document.Revisions.Count;
}
//Saves and closes the document
document.Save("Sample.docx", FormatType.Docx);
document.Close();
{% endhighlight %} 

{% endtabs %}

## How to enable track changes for Word document?

TrackChanges is used to keep track of the changes made to a Word document. This can be enabled by using the TrackChanges property of the Word document.

The following code example shows how to enable track changes of the document.

{% tabs %}   

{% highlight c# %}
//Creates a new Word document 
WordDocument document = new WordDocument();
//Adds new section to the document
IWSection section = document.AddSection();
//Adds new paragraph to the section
IWParagraph paragraph = section.AddParagraph();
//Appends text to the paragraph
IWTextRange text = paragraph.AppendText("This sample illustrates how to track the changes made to the word document. ");
//Sets font name and size for text
text.CharacterFormat.FontName = "Times New Roman";
text.CharacterFormat.FontSize = 14;
text = paragraph.AppendText("This track changes is useful in shared environment.");
text.CharacterFormat.FontSize = 12;
//Turns on the track changes option
document.TrackChanges = true;
//Saves and closes the document
document.Save("Sample.docx", FormatType.Docx);
document.Close();
{% endhighlight %} 

{% endtabs %}

