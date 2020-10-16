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

## How to set OpenType Font Features?

The Open type features provide special effects for the text. This feature is specific to Word 2010 and later version documents. The OpenType features includes the following:

* Ligatures – combination of characters, written as glyph
* Use Contextual Alternates – combination of letters based on surrounding characters
* Number spacing – specifies number width 
* Number forms – specifies number height
* Stylistic sets – specifies the look of the text, based on the font used

The following code illustrates how to set ligature types for text.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Add new text.
IWTextRange text = paragraph.appendText("Text to describe discretional ligatures");
//Set ligature type.
text.getCharacterFormat().setLigatures(LigatureType.Discretional);
text.getCharacterFormat().setFontName("Arial");
paragraph = section.addParagraph();
text = paragraph.appendText("Text to describe contextual ligatures");
text.getCharacterFormat().setLigatures(LigatureType.Contextual);
text.getCharacterFormat().setFontName("Arial");
paragraph = section.addParagraph();
text = paragraph.appendText("Text to describe historical ligatures");
text.getCharacterFormat().setLigatures(LigatureType.Historical);
text.getCharacterFormat().setFontName("Arial");
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to set contextual alternates.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Add new text.
IWTextRange text = paragraph.appendText("Text to describe contextual alternates");
text.getCharacterFormat().setFontName("Segoe Script");
//Set contextual alternates.
text.getCharacterFormat().setUseContextualAlternates(true);
paragraph = section.addParagraph();
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to set number spacing.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Add new text.
IWTextRange text = paragraph.appendText("Numbers to describe tabular number spacing 0123456789");
text.getCharacterFormat().setFontName("Calibri");
//Set number spacing.
text.getCharacterFormat().setNumberSpacing(NumberSpacingType.Tabular);
paragraph = section.addParagraph();
text = paragraph.appendText("Numbers to describe proportional number spacing 0123456789");
text.getCharacterFormat().setFontName("Calibri");
//Set number spacing.
text.getCharacterFormat().setNumberSpacing(NumberSpacingType.Proportional);
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to set number style.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Add new text.
IWTextRange text = paragraph.appendText("Numbers to describe oldstyle number form 0123456789");
text.getCharacterFormat().setFontName("Calibri");
//Set number style.
text.getCharacterFormat().setNumberForm(NumberFormType.OldStyle);
paragraph = section.addParagraph();
text = paragraph.appendText("Numbers to describe lining number form 0123456789");
text.getCharacterFormat().setFontName("Calibri");
//Set number style.
text.getCharacterFormat().setNumberForm(NumberFormType.Lining);
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to set different styles for the text.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add new section to the document.
IWSection section = document.addSection();
//Add new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Add new text.
IWTextRange text = paragraph.appendText("Text to describe stylistic sets");
text.getCharacterFormat().setFontName("Gabriola");
//Set stylistic set.
text.getCharacterFormat().setStylisticSet(StylisticSetType.StylisticSet06);
paragraph = section.addParagraph();
//Add new text.
text = paragraph.appendText("Text to describe stylistic sets");
text.getCharacterFormat().setFontName("Gabriola");
//Set stylistic set.
text.getCharacterFormat().setStylisticSet(StylisticSetType.StylisticSet15);
//Save and close the document.
document.save("Sample.docx", FormatType.Docx);
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

### Headers and Footers

The headers and footers can be inserted with text, graphics, and any other information that is contained in the document. 

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

### Tables

Tables are used to organize information and to display the information in rows and columns. You can also add images or even other tables to the table.

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

