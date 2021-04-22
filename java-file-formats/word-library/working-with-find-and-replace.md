---
title: Working with Find and Replace | Syncfusion
description: This section illustrates finding a text and replacing it with a new text in a Word document without Microsoft Word or Office interop
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Find and Replace

You can search a particular text you like to change and replace it with another text or part of the document.

## Finding contents in a Word document

You can find the first occurrence of a particular text within a single paragraph in the document by using `Find` method and its next occurrence by using `FindNext` method. You can also find a particular text pattern in the document.

The following code example illustrates how to find a particular text and its next occurrence in the document.

{% tabs %}  

{% highlight JAVA %}
//Loads the template document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Finds the first occurrence of a particular text in the document
TextSelection textSelection = document.find("as graphical contents", false, true);
//Gets the found text as single text range
WTextRange textRange = textSelection.getAsOneRange();
//Modifies the text
textRange.setText("Replaced text");
//Sets highlight color
textRange.getCharacterFormat().setHighlightColor(ColorSupport.getYellow());
//Finds the next occurrence of a particular text from the previous paragraph
textSelection = document.findNext(textRange.getOwnerParagraph(), "paragraph", true, false);
//Gets the found text as single text range
WTextRange range = textSelection.getAsOneRange();
//Sets bold formatting
range.getCharacterFormat().setBold(true);
//Saves and closes the document
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

You can find all the occurrence of a particular text within a single paragraph in the document by using `FindAll` method. 

The following code example illustrates how to find all the occurrences of a particular text in the document.

{% tabs %} 

{% highlight JAVA %}
//Loads the template document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Finds all the occurrences of a particular text
TextSelection[] textSelections = document.findAll("paragraph",false,true);
for(Object textSelection_tempObj : textSelections)
{
	TextSelection textSelection = (TextSelection)textSelection_tempObj;
	WTextRange textRange = textSelection.getAsOneRange();
	textRange.getCharacterFormat().setHighlightColor(ColorSupport.getYellowGreen());
}
//Saves and closes the document
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

You can find the first occurrence of a particular text extended to several paragraphs in the document by using `FindSingleLine` method and its next occurrence by using `FindNextSingleLine` method.

The following code example illustrates how to find a particular text extended to several paragraphs in the Word document.

{% tabs %}   

{% highlight JAVA %}
//Loads the template document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Finds the first occurrence of a particular text extended to several paragraphs in the document
TextSelection[] textSelections = document.findSingleLine("First paragraph Second paragraph", true, false);
WParagraph paragraph = null;
for(Object textSelection_tempObj : textSelections)
{
	//Gets the found text as single text range and set highlight color
	TextSelection textSelection = (TextSelection)textSelection_tempObj;
	WTextRange textRange = textSelection.getAsOneRange();
	textRange.getCharacterFormat().setHighlightColor(ColorSupport.getYellowGreen());
	paragraph=textRange.getOwnerParagraph();
}
//Finds the next occurrence of a particular text extended to several paragraphs in the document
textSelections=document.findNextSingleLine(paragraph,"First paragraph Second paragraph",true,false);
for(Object textSelection_tempObj : textSelections)
{
	//Gets the found text as single text range and sets italic formatting
	TextSelection textSelection = (TextSelection)textSelection_tempObj;
	WTextRange text = textSelection.getAsOneRange();
	text.getCharacterFormat().setItalic(true);
}
//Saves and closes the document
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %} 

## Find and replace text with other text

You can find text in a Word document and replace it with other text. Unlike the `find` method, the `replace` method replaces all occurrences of the text. You can customize it to replace only the first occurrence of a text by setting the `setReplaceFirst` property of the WordDocument class to true.

The following code example illustrates how to replace all occurrences of a misspelled word with the correctly spelled word.

{% tabs %}  

{% highlight JAVA %}
// Opens the input Word document
WordDocument document = new WordDocument("Template.docx");
// Finds all occurrences of a misspelled word and replaces with properly spelled word
document.replace("Cyles", "Cycles", true, true);
//Saves and closes the document
document.save("Sample.docx");
document.close();
{% endhighlight %}

{% endtabs %}

## Find and replace text with an image

You can find placeholder text in a Word document and replace it with any desired image.

The following code example illustrates how to find and replace text in a word document with an image

{% tabs %}  

{% highlight JAVA %}
//Opens the input Word document
WordDocument document = new WordDocument("Template.docx");
//Finds all the image placeholder text in the Word document
TextSelection[] textSelections = document.findAll(Pattern.compile(MatchSupport.trimPattern("^//(.*)")));
for (int i = 0; i < textSelections.length; i++) 
{
	// Replaces the image placeholder text with desired image
	WParagraph paragraph = new WParagraph(document);
	WPicture picture = (WPicture) paragraph.appendPicture(new FileInputStream(textSelections[i].getSelectedText() + ".png"));
	TextSelection newSelection = new TextSelection(paragraph, 0, 1);
	TextBodyPart bodyPart = new TextBodyPart(document);
	bodyPart.getBodyItems().add(paragraph);
	document.replace(textSelections[i].getSelectedText(), bodyPart, true, true);
}
//Saves and closes the document
document.save("Sample.docx");
document.close();
{% endhighlight %}

{% endtabs %}

## Find and replace a pattern of text with a merge field 

You can find and replace a pattern of text in a Word document with merge fields using Regex.

The following code example illustrates how to create a mail merge template by replacing a pattern of text (enclosed within ‘«’ and ‘»’) in a Word document with the desired merge fields.

{% tabs %}  

{% highlight JAVA %}
// Opens the input Word document
WordDocument document = new WordDocument("Template.docx");
// Finds all the placeholder text enclosed within '«' and '»' in the Word document
TextSelection[] textSelections = document.findAll(
Pattern.compile(MatchSupport.trimPattern("«([(?i)image(?-i)]*:*[a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»")));
String[] searchedPlaceholders = new String[textSelections.length];
for (int i = 0; i < textSelections.length; i++) 
{
	searchedPlaceholders[(int) i] = textSelections[(int) i].getSelectedText();
}
for (int i = 0; i < searchedPlaceholders.length; i++) 
{
	WParagraph paragraph = new WParagraph(document);
	// Replaces the placeholder text enclosed within '«' and '»' with desired merge field
	paragraph.appendField(StringSupport.trimEnd(StringSupport.trimStart(searchedPlaceholders[i], '«'), '»'),FieldType.FieldMergeField);
	TextSelection newSelection = new TextSelection(paragraph, 0, paragraph.getItems().getCount());
	TextBodyPart bodyPart = new TextBodyPart(document);
	bodyPart.getBodyItems().add(paragraph);
	document.replace(searchedPlaceholders[(int) i], bodyPart, true, true, true);
}
//Saves and closes the document
document.save("Sample.docx");
document.close();
{% endhighlight %}

{% endtabs %}

## Find and replace text with a table 

You can find placeholder text in a Word document and replace it with a table.

The following code example illustrates how to do this.

{% tabs %}  

{% highlight JAVA %}
// Opens the input Word document
WordDocument document = new WordDocument("Template.docx");
// Creates a new table
WTable table = new WTable(document);
table.resetCells(1, 6);
table.get(0, 0).setWidth(52f);
table.get(0, 0).addParagraph().appendText("Supplier ID");
table.get(0, 1).setWidth(128f);
table.get(0, 1).addParagraph().appendText("Company Name");
table.get(0, 2).setWidth(70f);
table.get(0, 2).addParagraph().appendText("Contact Name");
table.get(0, 3).setWidth(92f);
table.get(0, 3).addParagraph().appendText("Address");
table.get(0, 4).setWidth(66.5f);
table.get(0, 4).addParagraph().appendText("City");
table.get(0, 5).setWidth(56f);
table.get(0, 5).addParagraph().appendText("Country");
// Imports data to the table
importDataToTable(table);
// Applies the built-in table style (Medium Shading 1 Accent 1) to the table
table.applyStyle(BuiltinTableStyle.MediumShading1Accent1);
TextBodyPart bodyPart = new TextBodyPart(document);
bodyPart.getBodyItems().add(table);
// Replaces the table placeholder text with a new table
document.replace("[Suppliers table]", bodyPart, true, true, true);
// Saves and closes the document
document.save("Sample.docx");
{% endhighlight %}

{% endtabs %}

The following code example provides supporting method for the above code.

{% tabs %}  

{% highlight JAVA %}
private void importDataToTable(WTable table) throws Exception 
{
	FileStreamSupport fs = new FileStreamSupport("Suppliers.xml", FileMode.Open, FileAccess.Read);
	XmlReaderSupport reader = XmlReaderSupport.create(fs);
	if (reader == null)
		throw new Exception("reader");
	while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
		reader.read();
	if (reader.getLocalName() != "SuppliersList")
		throw new Exception(StringSupport.concat("Unexpected xml tag ", reader.getLocalName()));
	reader.read();
	while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
	reader.read();
	while (reader.getLocalName() != "SuppliersList") 
	{
		if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) 
		{
			switch (reader.getLocalName()) 
			{
				case "Suppliers":
				WTableRow tableRow = table.addRow(true);
				importDataToRow(reader, tableRow);
				break;
			}
		} 
		else 
		{
			reader.read();
			if ((reader.getLocalName() == "SuppliersList") && reader.getNodeType() == XmlNodeType.EndElement)
				break;
		}
	}
	reader.close();
	fs.close();
}
{% endhighlight %}

{% endtabs %}

The following code example provides supporting method for the above code.

{% tabs %}  

{% highlight JAVA %}
private void importDataToRow(XmlReaderSupport reader, WTableRow tableRow) throws Exception 
{
	if (reader == null)
		throw new Exception("reader");
	while (reader.getNodeType().getEnumValue() != XmlNodeType.Element.getEnumValue())
		reader.read();
	if (reader.getLocalName() != "Suppliers")
		throw new Exception(StringSupport.concat("Unexpected xml tag ", reader.getLocalName()));
	reader.read();
	while (reader.getNodeType().getEnumValue() == XmlNodeType.Whitespace.getEnumValue())
		reader.read();
	while (reader.getLocalName() != "Suppliers") 
	{
		if (reader.getNodeType().getEnumValue() == XmlNodeType.Element.getEnumValue()) 
		{
			switch (reader.getLocalName()) 
			{
				case "SupplierID":
					tableRow.getCells().get(0).addParagraph().appendText(reader.readContentAsString());
					break;
				case "CompanyName":
					tableRow.getCells().get(1).addParagraph().appendText(reader.readContentAsString());
					break;
				case "ContactName":
					tableRow.getCells().get(2).addParagraph().appendText(reader.readContentAsString());
					break;
				case "Address":
					tableRow.getCells().get(3).addParagraph().appendText(reader.readContentAsString());
					break;
				case "City":
					tableRow.getCells().get(4).addParagraph().appendText(reader.readContentAsString());
					break;
				case "Country":
					tableRow.getCells().get(5).addParagraph().appendText(reader.readContentAsString());
					break;
				default:
					reader.skip();
					break;
			}
		} 
		else 
		{
			reader.read();
			if ((reader.getLocalName() == "Suppliers") && reader.getNodeType() == XmlNodeType.EndElement)
				break;
		}
	}
}
{% endhighlight %}

{% endtabs %}

## Find and replace text in Word document with another document 

You can find and replace text with another Word document.

The following code example illustrates how to merge or combine Word documents by replacing text with another document (the content of a subheading).

{% tabs %}  

{% highlight JAVA %}
// Opens the Word template document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
TextSelection[] textSelections = document.findAll(Pattern.compile(MatchSupport.trimPattern("\\[(.*)\\]")));
for (int i = 0; i < textSelections.length; i++) 
{
	WordDocument subDocument = new WordDocument(StringSupport.trimEnd(StringSupport.trimStart(textSelections[i].getSelectedText(), '['), ']') + ".docx",FormatType.Docx);
	document.replace(textSelections[(int) i].getSelectedText(), subDocument, true, true);
	subDocument.close();
}
// Saves the Word document
document.save("Sample.docx");
// Closes the document
document.close();
{% endhighlight %}

{% endtabs %}
## Find and replace text extending to several paragraphs

Apart from finding text in a paragraph, you can also find and replace text that extends to several paragraphs in a Word document. You can find the first occurrence of the text that extends to several paragraphs by using the `findSingleLine` method. Find the next occurrences of the text by using the `findNextSingleLine` method. Similarly, you can replace text that extends to several paragraphs by using `replaceSingleLine` method.

The following code example illustrates how to replace text that extends to several paragraphs.

{% tabs %}  

{% highlight JAVA %}
// Opens the input Word document
WordDocument document = new WordDocument("Template.docx",FormatType.Docx);
WordDocument subDocument = new WordDocument("Source.docx", FormatType.Docx);
// Gets the content from another Word document to replace
TextBodyPart replacePart = new TextBodyPart(subDocument);
for (Object bodyItem_tempObj : subDocument.getLastSection().getBody().getChildEntities()) 
{
	TextBodyItem bodyItem = (TextBodyItem) bodyItem_tempObj;
	replacePart.getBodyItems().add(bodyItem.clone());
}
String placeholderText = "Suppliers/Vendors of Northwind" + "Customers of Northwind" + "Employee details of Northwind traders" + "The product information" + "The inventory details" + "The shippers" + "Purchase Order transactions" + "Sales Order transaction" + "Inventory transactions" + "Invoices" + "[end replace]";
// Finds the text that extends to several paragraphs and replaces it with desired content.
document.replaceSingleLine(placeholderText, replacePart, false, false);
subDocument.close();
// Saves the Word document
document.save("Sample.docx");
// Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

## Find text in a Word document and format 

You can find text in a Word document and format or highlight it .You can find the first occurrence of text using the `find` method. Find the next occurrences of the text using the `findNext` method.

The following code example illustrates how to find all occurrences of a length of text and highlight it.

{% tabs %}  

{% highlight JAVA %}
// Opens the input Word document
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Finds all occurrence of the text in the Word document
TextSelection[] textSelections = document.findAll("Adventure", true, true);
for (int i = 0; i < textSelections.length; i++) 
{
	// Sets the highlight color for the searched text as Yellow
	textSelections[(int) i].getAsOneRange().getCharacterFormat().setHighlightColor(ColorSupport.getYellow());
}
// Saves the Word document
document.save("Sample.docx");
// Closes the document
document.close();
{% endhighlight %}

{% endtabs %}
