---
title: Working with Content Controls | Word library | Syncfusion
description: This section illustrates how to work with Content Controls in Word document using Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Working with Content Controls

## What is Content Control?

Content controls are individual controls that you can add and customize to use in templates, forms, and documents. For example, many online forms are designed with a drop-down list control that provides a restricted set of choices.

N> You can use content controls only in documents that are saved in the Open XML Format.

Content controls can be categorized based on its occurrence in a document as follows,

* InlineContentControl: Among inline content inside, as a child of a paragraph.
* BlockContentControl: Among paragraphs and tables, as a child of a Body, HeaderFooter, Comment, Footnote, or a Shape node.

### Block Content Control

You can add content control to a text body of the Word document using block content control. You can add text, tables, pictures, or other items into the block content control. Refer to the following code.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds new section to the document.
IWSection section = document.addSection();
WTextBody textBody = section.getBody();
//Adds block content control into Word document.
BlockContentControl blockContentControl = (BlockContentControl)textBody.addBlockContentControl(ContentControlType.RichText);
//Adds new paragraph in the block content control.
WParagraph paragraph = (WParagraph)blockContentControl.getTextBody().addParagraph();
//Adds new text to the paragraph.
paragraph.appendText("Block content control");
//Adds new table to the block content control.
WTable table = (WTable)blockContentControl.getTextBody().addTable();
//Specifies the total number of rows and columns.
table.resetCells(2,3);
paragraph = (WParagraph)blockContentControl.getTextBody().addParagraph();
//Adds image to the paragraph.
paragraph.appendPicture(new FileInputStream("Image.png"));
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Inline Content Control

You can add content control as a child to a paragraph using the inline content control. You can add text, pictures, fields or other paragraph items into the inline content control. Refer to the following code.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends inline content control to the paragraph.
InlineContentControl inlineContentControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.RichText);
WTextRange textRange = new WTextRange(document);
//Adds new text to the inline content control.
textRange.setText("Inline content control ");
inlineContentControl.getParagraphItems().add(textRange);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

N> Currently, DocIO does not support RowContentControl and CellContentControl.

## Common properties of Content Control

You can set formatting options for the content control in the Word document. The following are the common properties of a content control.

### Title

The title of the content control. 

### Tag

The tag value to identify the content control.

### Appearance

This property allows you to define the appearance of the content controls. The appearance can be any one of the following:

* BoundingBox: Displays the contents of content control within a box.
* Tags: Displays the contents of content control within tags.
* Hidden: Displays the contents of content control without any box or tags.

### Color

Defines the color of the content control.

### Temporary 

This property defines whether to remove a content control from the Word document when you edit the contents of the content control.

### Lock Contents

Locking the contents of the content control. It restricts to modify the contents of the content control.

### Lock Content Control

It restricts to remove or delete the content control.

### Example – Content Control Common properties

The following code sample illustrates the content control properties usage.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends rich text content control to the paragraph.
IInlineContentControl contentControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.RichText);
WTextRange textRange = new WTextRange(document);
textRange.setText("Rich text content control.");
//Adds new text to the rich text content control.
contentControl.getParagraphItems().add(textRange);
//Sets tag appearance for the content control.
contentControl.getContentControlProperties().setAppearance(ContentControlAppearance.Tags);
//Sets a tag property to identify the content control.
contentControl.getContentControlProperties().setTag("Rich Text");
//Sets a title for the content control.
contentControl.getContentControlProperties().setTitle("Text");
//Sets the color for the content control.
contentControl.getContentControlProperties().setColor(ColorSupport.getMagenta());
//Gets the type of content control.
ContentControlType controlType = contentControl.getContentControlProperties().getType();
//Enables content control lock.
contentControl.getContentControlProperties().setLockContentControl(true);
//Protects the contents of content control.
contentControl.getContentControlProperties().setLockContents(true);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

## Why Content Control?

The content controls have the following three major use cases:

* Protection
* Form Filling
* Data Binding with Content Controls (XML Mapping)

### Protection

Content controls provides options to prevent users from editing or deleting parts of a Word document contents. This is useful if you have information in a Word document or template that you should be able to read but not edit, or if you want to be able to edit content controls but not delete them. 

To protect contents inside a content control, you can use properties of the content control to prevent editing or deleting the content control:

* The **LockContents** property prevents from editing the contents of the content control.
* The **LockContentControl** property prevents from deleting the content control.

The following code sample shows how to protect the content control and its contents.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document. 
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends rich text content control to the paragraph.
IInlineContentControl contentControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.RichText);
WTextRange textRange = new WTextRange(document);
textRange.setText("Rich text content control.");
//Adds new text to the rich text content control.
contentControl.getParagraphItems().add(textRange);
//Sets tag appearance for the content control.
contentControl.getContentControlProperties().setAppearance(ContentControlAppearance.Tags);
//Sets a tag property to identify the content control.
contentControl.getContentControlProperties().setTag("Rich Text Protected");
//Sets a title for the content control.
contentControl.getContentControlProperties().setTitle("Text Protected");
//Enables content control lock.
contentControl.getContentControlProperties().setLockContentControl(true);
//Protects the contents of content control.
contentControl.getContentControlProperties().setLockContents(true);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Form Filling

Another major use case is to create the forms. You can design your own forms for various stages using the text box, check box, and list box. Refer to the following code example. 

Form creation:

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document .
WordDocument document = new WordDocument();
//Adding a new section to the document.
IWSection section = document.addSection();
//Adding a new paragraph to the section.

//Document formatting.
//Sets background color for document.
IWParagraph paragraph = section.addParagraph();
document.getBackground().getGradient().setColor1(ColorSupport.fromArgb(232,232,232));
document.getBackground().getGradient().setColor2(ColorSupport.fromArgb(255,255,255));
document.getBackground().setType(BackgroundType.Gradient);
document.getBackground().getGradient().setShadingStyle(GradientShadingStyle.Horizontal);
document.getBackground().getGradient().setShadingVariant(GradientShadingVariant.ShadingDown);
//Sets page size for document.
section.getPageSetup().getMargins().setAll(30f);
section.getPageSetup().setPageSize(new SizeFSupport(600,600f));

//Title Section.
//Adds a new table to the section.
IWTable table = section.getBody().addTable();
table.resetCells(1,2);
//Gets the table first row.
WTableRow row = table.getRows().get(0);
row.setHeight(25f);
//Adds a new paragraph to the cell.
IWParagraph cellPara = row.getCells().get(0).addParagraph();
//Appends a new picture.
IWPicture pic = cellPara.appendPicture(new FileInputStream("Image.jpg"));
pic.setHeight((float)80);
pic.setWidth((float)180);
//Adds a new paragraph to the next cell.
cellPara = row.getCells().get(1).addParagraph();
row.getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
row.getCells().get(1).getCellFormat().setBackColor(ColorSupport.fromArgb(173,215,255));
//Appends the text.
IWTextRange txt = cellPara.appendText("Job Application Form");
cellPara.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//Sets the formats.
txt.getCharacterFormat().setBold(true);
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(18f);
//Sets the width and border type.
row.getCells().get(0).setWidth((float)200);
row.getCells().get(1).setWidth((float)300);
row.getCells().get(1).getCellFormat().getBorders().setBorderType(BorderStyle.Hairline);
//Adds a new paragraph.
section.addParagraph();

//General Information.
//Adds a new table.
table=section.getBody().addTable();
table.resetCells(2,1);
row = table.getRows().get(0);
row.setHeight((float)20);
row.getCells().get(0).setWidth((float)500);
//Adds a new paragraph.
cellPara = row.getCells().get(0).addParagraph();
//Sets a border type, color, background, and vertical alignment for cell.
row.getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Thick);
row.getCells().get(0).getCellFormat().getBorders().setColor(ColorSupport.fromArgb(155,205,255));
row.getCells().get(0).getCellFormat().setBackColor(ColorSupport.fromArgb(198,227,255));
row.getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
txt=cellPara.appendText(" General Information");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setBold(true);
txt.getCharacterFormat().setFontSize(11f);
row = table.getRows().get(1);
cellPara=row.getCells().get(0).addParagraph();
//Sets a width, border type, color and background for cell
row.getCells().get(0).setWidth((float)500);
row.getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Hairline);
row.getCells().get(0).getCellFormat().getBorders().setColor(ColorSupport.fromArgb(155,205,255));
row.getCells().get(0).getCellFormat().setBackColor(ColorSupport.fromArgb(222,239,255));
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n Full Name:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
InlineContentControl txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Text);
txtField.getContentControlProperties().setTitle("Text");
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n Birth Date:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Date);
txtField.getContentControlProperties().setTitle("Date");
//Sets the date display format.
txtField.getContentControlProperties().setDateDisplayFormat("M/d/yyyy");
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt=cellPara.appendText("\n\n Address:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Text);
txtField.getContentControlProperties().setTitle("Text");
//Sets multiline property to true to get the multiple line input of Address.
txtField.getContentControlProperties().setMultiline(true);
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n Phone:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Text);
txtField.getContentControlProperties().setTitle("Text");
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n Email:\t\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Text);
txtField.getContentControlProperties().setTitle("Text");
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
cellPara.appendText("\n");
section.addParagraph();

//Educational Qualification.
//Adds a new table to the section.
table = section.getBody().addTable();
table.resetCells(2,1);
row = table.getRows().get(0);
row.setHeight((float)20);
//Sets width, border type, color, background and vertical alignment for cell.
row.getCells().get(0).setWidth((float)500);
row.getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Thick);
row.getCells().get(0).getCellFormat().getBorders().setColor(ColorSupport.fromArgb(155,205,255));
row.getCells().get(0).getCellFormat().setBackColor(ColorSupport.fromArgb(198,227,255));
row.getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
cellPara=row.getCells().get(0).addParagraph();
//Appends a text to paragraph of cell.
txt = cellPara.appendText(" Educational Qualification");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setBold(true);
txt.getCharacterFormat().setFontSize(11f);
row = table.getRows().get(1);
//Sets width, border type, color, and background for cell.
row.getCells().get(0).setWidth((float)500);
row.getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Hairline);
row.getCells().get(0).getCellFormat().getBorders().setColor(ColorSupport.fromArgb(155,205,255));
row.getCells().get(0).getCellFormat().setBackColor(ColorSupport.fromArgb(222,239,255));
cellPara = row.getCells().get(0).addParagraph();
//Appends a text to paragraph of cell.
txt=cellPara.appendText("\n Type:\t\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
InlineContentControl dropdown = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.DropDownList);
WTextRange textRange = new WTextRange(document);
textRange.setText("Choose an item from drop down list");
dropdown.getParagraphItems().add(textRange);
//Creates an item for dropdown list.
ContentControlListItem item = new ContentControlListItem();
//Sets the text to be displayed as list item.
item.setDisplayText("Higher");
//Sets the value to the list item.
item.setValue("1");
//Adds item to the dropdown list.
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Vocational");
item.setValue("2");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Universal");
item.setValue("3");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
dropdown.getContentControlProperties().setTitle("Drop-Down");
//Sets formatting options for text present inside a content control.
dropdown.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
dropdown.getBreakCharacterFormat().setFontName("Arial");
dropdown.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n Institution:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a new inline content control to enter the value.
txtField = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.Text);
//Sets formatting options for text present inside a content control.
txtField.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
txtField.getBreakCharacterFormat().setFontName("Arial");
txtField.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n Programming Languages:");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n\t C#:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(9f);
//Appends a new inline content control to enter the value.
dropdown = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.DropDownList);
textRange = new WTextRange(document);
textRange.setText("Choose an item from drop down list");
dropdown.getParagraphItems().add(textRange);
//Creates an item for dropdown list.
item = new ContentControlListItem();
item.setDisplayText("Perfect");
//Sets the value to the list item.
item.setValue("1");
//Adds item to the dropdown list.
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Good");
item.setValue("2");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Excellent");
item.setValue("3");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
//Sets formatting options for text present inside a content control.
dropdown.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
dropdown.getBreakCharacterFormat().setFontName("Arial");
dropdown.getBreakCharacterFormat().setFontSize(11f);
//Appends a text to paragraph of cell.
txt = cellPara.appendText("\n\n\t VB:\t\t\t\t");
txt.getCharacterFormat().setFontName("Arial");
txt.getCharacterFormat().setFontSize(9f);
//Appends a new inline content control to enter the value.
dropdown = (InlineContentControl)cellPara.appendInlineContentControl(ContentControlType.DropDownList);
textRange = new WTextRange(document);
textRange.setText("Choose an item from drop down list");
dropdown.getParagraphItems().add(textRange);
//Creates an item for dropdown list.
item = new ContentControlListItem();
//Sets the text to be displayed as list item.
item.setDisplayText("Perfect");
//Sets the value to the list item.
item.setValue("1");
//Adds item to the dropdown list.
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Good");
item.setValue("2");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Excellent");
item.setValue("3");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
dropdown.getContentControlProperties().setTitle("Drop-Down");
//Sets formatting options for text present inside a content control.
dropdown.getBreakCharacterFormat().setTextColor(ColorSupport.getMidnightBlue());
dropdown.getBreakCharacterFormat().setFontName("Arial");
dropdown.getBreakCharacterFormat().setFontSize(11f);
//Saves and closes the Word document instance.
document.save("Form_Template.docx");
document.close();
{% endhighlight %}

{% endtabs %}

You can also fill the forms using the DocIO. Refer to the following code example.

Form filling:

{% tabs %}
{% highlight JAVA %}
//Open the created form document.
WordDocument document1 = new WordDocument("Form_Template.docx");
IWSection sec = document1.getLastSection();
InlineContentControl inlineCC;
InlineContentControl dropDownCC;
WTable table1 = (WTable)sec.getTables().get(1);
WTableRow row1 = table1.getRows().get(1);

//General Information.
//Fill the name
WParagraph cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(1);
inlineCC = (InlineContentControl)cellPara1.getChildEntities().getLastItem();
WTextRange text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("Steve Jobs");
inlineCC.getParagraphItems().add(text);
//Fill the date of birth.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(3);
inlineCC = (InlineContentControl)cellPara1.getChildEntities().getLastItem();
text=new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("06/01/1994");
inlineCC.getParagraphItems().add(text);
//Fill the address.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(5);
inlineCC=(InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("2501 Aerial Center Parkway.");
inlineCC.getParagraphItems().add(text);
text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("Morrisville, NC 27560.");
inlineCC.getParagraphItems().add(text);
text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("USA.");
inlineCC.getParagraphItems().add(text);
//Fill the phone no.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(7);
inlineCC = (InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("+1 919.481.1974");
inlineCC.getParagraphItems().add(text);
//Fill the email id.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(9);
inlineCC=(InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(inlineCC.getBreakCharacterFormat());
text.setText("steve123@email.com");
inlineCC.getParagraphItems().add(text);

//Educational Information.
table1=(WTable)sec.getTables().get(2);
row1 = table1.getRows().get(1);
//Fill the education type.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(1);
dropDownCC = (InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(dropDownCC.getBreakCharacterFormat());
text.setText(dropDownCC.getContentControlProperties().getContentControlListItems().get(1).getDisplayText());
dropDownCC.getParagraphItems().add(text);
//Fill the university.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(3);
inlineCC = (InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(dropDownCC.getBreakCharacterFormat());
text.setText("Michigan University");
inlineCC.getParagraphItems().add(text);
//Fill the C# experience level.
cellPara1 = (WParagraph)row1.getCells().get(0).getChildEntities().get(7);
dropDownCC=(InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(dropDownCC.getBreakCharacterFormat());
text.setText(dropDownCC.getContentControlProperties().getContentControlListItems().get(2).getDisplayText());
dropDownCC.getParagraphItems().add(text);
//Fill the VB experience level.
cellPara1=(WParagraph)row1.getCells().get(0).getChildEntities().get(9);
dropDownCC=(InlineContentControl)cellPara1.getChildEntities().getLastItem();
text = new WTextRange(document1);
text.applyCharacterFormat(dropDownCC.getBreakCharacterFormat());
text.setText(dropDownCC.getContentControlProperties().getContentControlListItems().get(1).getDisplayText());
dropDownCC.getParagraphItems().add(text);
//Saves and closes the Word document instance.
document1.save("Form_Filled.docx");
document1.close();
{% endhighlight %}

{% endtabs %}

## Types of Content Controls

The following types of content controls can be created by using the Essential DocIO.

* Rich Text
* Plain Text
* Check Box
* Drop-Down List and Combo Box
* Picture

### Rich Text

A rich text content control contains text or other items, such as tables, pictures, or other content controls. The following code illustrates how to add new rich text content control. 

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
InlineContentControl richTextControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.RichText);
WTextRange textRange = new WTextRange(document);
textRange.setText("Rich text content control.");
//Adds new text to the rich text content control.
richTextControl.getParagraphItems().add(textRange);
WPicture picture = new WPicture(document);
picture.loadImage(new FileInputStream("Image.png"));
picture.setHeight((float)100);
picture.setWidth((float)100);
//Adds new picture to the rich text content control.
richTextControl.getParagraphItems().add(picture);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Plain Text

A plain text content control contains text and cannot contain other items, such as tables, pictures, or other content controls. Refer to the following code to add plain text content control.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends plain text content control to the paragraph.
InlineContentControl plainTextControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.Text);
WTextRange textRange = new WTextRange(document);
textRange.setText("Plain text content control.");
//Adds new text to the plain text content control.
plainTextControl.getParagraphItems().add(textRange);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Check Box

A check box content control provides a UI that represents a binary state: checked or unchecked. Default state for check box is unchecked. Refer to the following code to add check box content control.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document .
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends checkbox content control to the paragraph.
InlineContentControl checkBox = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.CheckBox);
checkBox.getContentControlProperties().setIsChecked(true);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Drop-Down List and Combo Box

A drop-down list content control and combo box content control displays a list of items you can select. Unlike a drop-down list, the combo box allows to add your own items. Refer to the following code to add drop-down list and combo box content controls.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds new section to the document.
IWSection section = document.addSection();
//Adds new paragraph to the section.
WParagraph paragraph = (WParagraph)section.addParagraph();
InlineContentControl dropdown = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.DropDownList);
WTextRange textRange = new WTextRange(document);
//Sets default option to display.
textRange.setText("Choose an item from drop down list");
dropdown.getParagraphItems().add(textRange);
//Creates an item for dropdown list.
ContentControlListItem item = new ContentControlListItem();
//Sets the text to be displayed as list item.
item.setDisplayText("ASP.NET MVC");
//Sets the value to the list item.
item.setValue("1");
//Adds item to the dropdown list.
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Windows Forms");
item.setValue("2");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("WPF");
item.setValue("3");
dropdown.getContentControlProperties().getContentControlListItems().add(item);
//Adds new paragraph to the section.
paragraph = (WParagraph)section.addParagraph();
//Appends combo box content control to the paragraph.
InlineContentControl comboBox = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.ComboBox);
textRange=new WTextRange(document);
//Sets default option to display.
textRange.setText("Choose an item from combo box");
comboBox.getParagraphItems().add(textRange);
//Creates an item for combo box.
item = new ContentControlListItem();
//Sets the text to be displayed as list item.
item.setDisplayText("Word to HTML");
//Sets the value to the list item.
item.setValue("1");
comboBox.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Word to Image");
item.setValue("2");
comboBox.getContentControlProperties().getContentControlListItems().add(item);
item = new ContentControlListItem();
item.setDisplayText("Word to PDF");
item.setValue("3");
comboBox.getContentControlProperties().getContentControlListItems().add(item);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Picture

A picture content control displays an image. Refer to the following code to add new picture content control.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document.
document.ensureMinimal();
//Gets the last paragraph.
WParagraph paragraph = document.getLastParagraph();
//Adds text to the paragraph.
paragraph.appendText("A new text is added to the paragraph. ");
//Appends picture content control to the paragraph.
InlineContentControl pictureContentControl = (InlineContentControl)paragraph.appendInlineContentControl(ContentControlType.Picture);
//Creates a new image instance and load image.
WPicture picture = new WPicture(document);
picture.loadImage(new FileInputStream("Image.png"));
//Adds picture to the picture content control.
pictureContentControl.getParagraphItems().add(picture);
//Saves and closes the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}