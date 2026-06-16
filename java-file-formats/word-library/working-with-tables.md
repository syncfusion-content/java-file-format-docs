---
title: Working with Shapes | Syncfusion
description: This section describes how to work with the shapes and group shapes in a Word document using the Syncfusion Java Word library (Essential DocIO).
platform: java-file-formats
control: Word Library
documentation: UG
keywords: 
---
# Working with Shapes in a Word Document

Shapes are drawing objects that include lines, curves, circles, rectangles, etc. They can be preset or custom geometry. You can create and manipulate predefined shapes in DOCX and WordML format documents.

## Adding Shapes

The following code example shows how to add a predefined shape to the document.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.addSection();
// Add a new paragraph to the section.
WParagraph paragraph = (WParagraph) section.addParagraph();
// Add a new shape to the document.
Shape rectangle = paragraph.appendShape(AutoShapeType.RoundedRectangle, 150, 100);
// Set position for the shape.
rectangle.setVerticalPosition(72);
rectangle.setHorizontalPosition(72);
paragraph = (WParagraph) section.addParagraph();
// Add text body contents to the shape.
paragraph = (WParagraph) rectangle.getTextBody().addParagraph();
IWTextRange text = paragraph.appendText("This text is in a rounded rectangle shape");
text.getCharacterFormat().setTextColor(ColorSupport.getGreen());
text.getCharacterFormat().setBold(true);
// Add another shape to the document.
paragraph = (WParagraph) section.addParagraph();
paragraph.appendBreak(BreakType.LineBreak);
Shape pentagon = paragraph.appendShape(AutoShapeType.Pentagon, 100, 100);
paragraph = (WParagraph) pentagon.getTextBody().addParagraph();
paragraph.appendText("This text is in a pentagon shape");
// Set position for the shape.
pentagon.setHorizontalPosition(72);
pentagon.setVerticalPosition(200);
// Save and close the Word document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Format Shapes

Shapes can have formatting such as line color, fill color, positioning, wrap formats, etc. The following code example illustrates how to apply formatting options for a shape.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.addSection();
// Add a new paragraph to the section.
IWParagraph paragraph = (WParagraph) section.addParagraph();
// Append a shape to the paragraph.
Shape rectangle = paragraph.appendShape(AutoShapeType.RoundedRectangle, 150, 100);
rectangle.setVerticalPosition(72);
rectangle.setHorizontalPosition(72);
paragraph = (WParagraph) section.addParagraph();
paragraph = (WParagraph) rectangle.getTextBody().addParagraph();
IWTextRange text = paragraph.appendText("This text is in a rounded rectangle shape");
// Apply format to the text.
text.getCharacterFormat().setTextColor(ColorSupport.getGreen());
text.getCharacterFormat().setBold(true);
// Apply fill color for the shape.
rectangle.getFillFormat().setFill(true);
rectangle.getFillFormat().setColor(ColorSupport.getLightGray());
// Apply wrap formats.
rectangle.getWrapFormat().setTextWrappingStyle(TextWrappingStyle.Square);
rectangle.getWrapFormat().setTextWrappingType(TextWrappingType.Right);
// Set horizontal and vertical origin.
rectangle.setHorizontalOrigin(HorizontalOrigin.Margin);
rectangle.setVerticalOrigin(VerticalOrigin.Page);
// Set line format.
rectangle.getLineFormat().setDashStyle(LineDashing.Dot);
rectangle.getLineFormat().setColor(ColorSupport.getDarkGray());
// Save and close the Word document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Rotate Shapes

You can rotate a shape and also apply flipping (horizontal and vertical) to it. The following code example explains how to rotate and flip a shape.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.addSection();
// Add a new paragraph to the section.
WParagraph paragraph = (WParagraph) section.addParagraph();
Shape rectangle = paragraph.appendShape(AutoShapeType.RoundedRectangle, 150, 100);
// Set position for the shape.
rectangle.setVerticalPosition(72);
rectangle.setHorizontalPosition(72);
// Set 90-degree rotation.
rectangle.setRotation(90);
// Set horizontal flip.
rectangle.setFlipHorizontal(true);
paragraph = (WParagraph) section.addParagraph();
paragraph = (WParagraph) rectangle.getTextBody().addParagraph();
IWTextRange text = paragraph.appendText("This text is in a rounded rectangle shape");
// Save the Word document.
document.save("Result.docx", FormatType.Docx);
// Close the document.
document.close();
{% endhighlight %}

{% endtabs %}

## Grouping Shapes

The Word library now allows you to create or group multiple shapes, pictures, and text boxes as a group shape in a Word document (DOCX) and preserve it during DOCX and WordML format conversions.

You can create a document with group shapes by using Microsoft Word. It provides an option to group a set of shapes and images as a single shape or treat a group shape as an individual item.
![Create Group Shape in Microsoft Word](Working-with-Shapes_images/Working-with-Shapes_img1.jpeg)

**Key Features:**

1. You can easily manage a group of shapes, pictures, and text boxes as a group shape.
2. You can move several shapes or images simultaneously and apply the same formatting properties for children of group shapes.

N> 1. While grouping the shapes or other objects, the shapes should be positioned relative to the “Page”.
N> 2. While grouping the shapes or other objects, the wrapping style should not be "In Line with Text".

The following code example shows how to create a group shape in a Word document.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.addSection();
// Add a new paragraph to the section.
WParagraph paragraph = (WParagraph) section.addParagraph();
// Create a new group shape.
GroupShape groupShape = new GroupShape(document);
// Add the group shape to the paragraph.
paragraph.getChildEntities().add(groupShape);
// Create a new shape.
Shape shape = new Shape(document, AutoShapeType.RoundedRectangle);
// Set height and width for the shape.
shape.setHeight(100);
shape.setWidth(150);
// Set horizontal and vertical position.
shape.setHorizontalPosition(72);
shape.setVerticalPosition(72);
// Set wrapping style for the shape.
shape.getWrapFormat().setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
// Set horizontal and vertical origin.
shape.setHorizontalOrigin(HorizontalOrigin.Page);
shape.setVerticalOrigin(VerticalOrigin.Page);
// Add the specified shape to the group shape.
groupShape.add(shape);
// Create a new picture.
WPicture picture = new WPicture(document);
FileStreamSupport imageStream = new FileStreamSupport("Image.png", FileMode.Open, FileAccess.ReadWrite);
picture.loadImage(imageStream.toArray());
// Set wrapping style for the picture.
picture.setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
// Set height and width for the image.
picture.setHeight(100);
picture.setWidth(100);
// Set horizontal and vertical position.
picture.setHorizontalPosition(400);
picture.setVerticalPosition(150);
// Set horizontal and vertical origin.
picture.setHorizontalOrigin(HorizontalOrigin.Page);
picture.setVerticalOrigin(VerticalOrigin.Page);
// Add the specified picture to the group shape.
groupShape.add(picture);
// Create a new textbox.
WTextBox textbox = new WTextBox(document);
textbox.getTextBoxFormat().setWidth(150);
textbox.getTextBoxFormat().setHeight(75);
// Add new text to the textbox body.
IWParagraph textboxParagraph = textbox.getTextBoxBody().addParagraph();
textboxParagraph.appendText("Text inside text box");
// Set wrapping style for the textbox.
textbox.getTextBoxFormat().setTextWrappingStyle(TextWrappingStyle.Behind);
// Set horizontal and vertical position.
textbox.getTextBoxFormat().setHorizontalPosition(200);
textbox.getTextBoxFormat().setVerticalPosition(200);
// Set horizontal and vertical origin.
textbox.getTextBoxFormat().setVerticalOrigin(VerticalOrigin.Page);
textbox.getTextBoxFormat().setHorizontalOrigin(HorizontalOrigin.Page);
// Add the specified textbox to the group shape.
groupShape.add(textbox);
// Save and close the Word document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Nested Group Shapes

					   

The following code example illustrates how to group nested group shapes as a group shape in a Word document.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.addSection();
// Add a new paragraph to the section.
WParagraph paragraph = (WParagraph) section.addParagraph();
// Create a new group shape.
GroupShape groupShape = new GroupShape(document);
// Add the group shape to the paragraph.
paragraph.getChildEntities().add(groupShape);
// Append a new shape to the document.
Shape shape = new Shape(document, AutoShapeType.RoundedRectangle);
// Set height and width for the shape.
shape.setHeight(100);
shape.setWidth(150);
// Set wrapping style for the shape.
shape.getWrapFormat().setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
// Set horizontal and vertical position for the shape.
shape.setHorizontalPosition(72);
shape.setVerticalPosition(72);
// Set horizontal and vertical origin for the shape.
shape.setHorizontalOrigin(HorizontalOrigin.Page);
shape.setVerticalOrigin(VerticalOrigin.Page);
// Add the specified shape to the group shape.
groupShape.add(shape);
// Append a new picture to the document.
WPicture picture = new WPicture(document);
// Load image from the file.
FileStreamSupport imageStream = new FileStreamSupport("Image.png", FileMode.Open, FileAccess.ReadWrite);
picture.loadImage(imageStream.toArray());
// Set wrapping style for the picture.
picture.setTextWrappingStyle(TextWrappingStyle.InFrontOfText);
// Set height and width for the picture.
picture.setHeight(100);
picture.setWidth(100);
// Set horizontal and vertical position for the picture.
picture.setHorizontalPosition(400);
picture.setVerticalPosition(150);
// Set horizontal and vertical origin for the picture.
picture.setHorizontalOrigin(HorizontalOrigin.Page);
picture.setVerticalOrigin(VerticalOrigin.Page);
// Add the specified picture to the group shape.
groupShape.add(picture);
// Create a new nested group shape.
GroupShape nestedGroupShape = new GroupShape(document);
// Append a new textbox to the document.
WTextBox textbox = new WTextBox(document);
// Set width and height for the textbox.
textbox.getTextBoxFormat().setWidth(150);
textbox.getTextBoxFormat().setHeight(75);
// Add new text to the textbox body.
IWParagraph textboxParagraph = textbox.getTextBoxBody().addParagraph();
// Add new text to the textbox paragraph.
textboxParagraph.appendText("Text inside text box");
// Set wrapping style for the textbox.
textbox.getTextBoxFormat().setTextWrappingStyle(TextWrappingStyle.Behind);
// Set horizontal and vertical position for the textbox.
textbox.getTextBoxFormat().setHorizontalPosition(200);
textbox.getTextBoxFormat().setVerticalPosition(200);
// Set horizontal and vertical origin for the textbox.
textbox.getTextBoxFormat().setVerticalOrigin(VerticalOrigin.Page);
textbox.getTextBoxFormat().setHorizontalOrigin(HorizontalOrigin.Page);
// Add the specified textbox to the nested group shape.
nestedGroupShape.add(textbox);
// Append a new shape to the document.
shape = new Shape(document, AutoShapeType.Oval);
// Set height and width for the new shape.
shape.setHeight(100);
shape.setWidth(150);
// Set horizontal and vertical position for the shape.
shape.setHorizontalPosition(200);
shape.setVerticalPosition(72);
// Set horizontal and vertical origin for the shape.
shape.setHorizontalOrigin(HorizontalOrigin.Page);
shape.setVerticalOrigin(VerticalOrigin.Page);
// Set horizontal and vertical position for the nested group shape.
nestedGroupShape.setHorizontalPosition(72);
nestedGroupShape.setVerticalPosition(72);
// Add the specified shape to the nested group shape.
nestedGroupShape.add(shape);
// Add the nested group shape to the group shape of the paragraph.
groupShape.add(nestedGroupShape);
// Save and close the Word document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

## Ungrouping Shapes

You can ungroup group shapes in the Word document to preserve each shape as an individual item.

The following code example shows how to ungroup a group shape in a Word document.

{% tabs %}

{% highlight JAVA %}
// Load the template document.
WordDocument document = new WordDocument("Template.docx", FormatType.Automatic);
// Get the last paragraph.
WParagraph lastParagraph = document.getLastParagraph();
// Iterate through the paragraph items to get the group shape.
for (int i = 0; i < lastParagraph.getChildEntities().getCount(); i++)
{
	if (lastParagraph.getChildEntities().get(i) instanceof GroupShape)
	{
		GroupShape groupShape = (GroupShape) lastParagraph.getChildEntities().get(i);
		// Ungroup the child shapes in the group shape.
		groupShape.ungroup();
		break;
	}
}
// Save and close the Word document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Working with Table Style

A table style defines a set of table, row, cell, and paragraph-level formatting that can be applied to a table. The `WTableStyle` instance represents table style in a Word document.

N>  Essential<sup style="font-size:70%">&reg;</sup> DocIO currently provides support for table styles in DOCX and WordML formats alone. The visual appearance is also preserved in Word-to-HTML conversion.

The following code example illustrates how to apply the built-in table styles to the table.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument("Table.docx", FormatType.Docx);
WSection section = document.getSections().get(0);
WTable table = section.getTables().get(0);
// Apply "LightShading" built-in style to table.
table.applyStyle(BuiltinTableStyle.LightShading);
// Save and close the document instance.
document.save("TableStyle.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Table Style Options

Once you have applied a table style, you can enable or disable the special formatting of the table. There are six options: first column, last column, banded rows, banded columns, header row, and last row.

The following code example illustrates how to enable and disable the special table formatting options of the table styles.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument("Table.docx", FormatType.Docx);
WSection section = document.getSections().get(0);
WTable table = section.getTables().get(0);
// Apply "LightShading" built-in style to table.
table.applyStyle(BuiltinTableStyle.LightShading);
// Enable special formatting for banded columns of the table.
table.setApplyStyleForBandedColumns(true);
// Enable special formatting for banded rows of the table.
table.setApplyStyleForBandedRows(true);
// Disable special formatting for the first column of the table.
table.setApplyStyleForFirstColumn(false);
// Enable special formatting for the header row of the table.
table.setApplyStyleForHeaderRow(true);
// Enable special formatting for the last column of the table.
table.setApplyStyleForLastColumn(true);
// Disable special formatting for the last row of the table.
table.setApplyStyleForLastRow(false);
// Save and close the document instance.
document.save("TableStyle.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Custom Table Style

The following code example illustrates how to apply a custom table style to a table.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument("Table.docx", FormatType.Docx);
WSection section = document.getSections().get(0);
WTable table = section.getTables().get(0);
// Add a new custom table style.
WTableStyle tableStyle = (WTableStyle) document.addTableStyle("CustomStyle");
// Apply formatting for the whole table.
tableStyle.getTableProperties().setRowStripe(1);
tableStyle.getTableProperties().setColumnStripe(1);
tableStyle.getTableProperties().getPaddings().setTop(0);
tableStyle.getTableProperties().getPaddings().setBottom(0);
tableStyle.getTableProperties().getPaddings().setLeft(5.4f);
tableStyle.getTableProperties().getPaddings().setRight(5.4f);
// Apply conditional formatting for the first row.
ConditionalFormattingStyle firstRowStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstRow);
firstRowStyle.getCharacterFormat().setBold(true);
firstRowStyle.getCharacterFormat().setTextColor(ColorSupport.fromArgb(255, 255, 255, 255));
firstRowStyle.getCellProperties().setBackColor(ColorSupport.getBlue());
// Apply conditional formatting for the first column.
ConditionalFormattingStyle firstColumnStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstColumn);
firstColumnStyle.getCharacterFormat().setBold(true);
// Apply conditional formatting for odd rows.
ConditionalFormattingStyle oddRowBandingStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.OddRowBanding);
oddRowBandingStyle.getCellProperties().setBackColor(ColorSupport.getWhiteSmoke());
// Apply the custom table style to the table.
table.applyStyle("CustomStyle");
// Save and close the document instance.
document.save("TableStyle.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

### Apply Base Style

Table styles can be based on other table styles as well. When applying a base style, the new style will inherit the values of the base style that are not explicitly redefined in the new style. You can apply a custom table style or a built-in table style as a base for the table style.

The following code example illustrates how to apply built-in and custom table styles as base styles for another custom table.

{% tabs %}

{% highlight JAVA %}
// Create a new Word document.
WordDocument document = new WordDocument();
// Add one section and paragraph to the document.
document.ensureMinimal();
WTable table = (WTable) document.getLastSection().addTable();
table.resetCells(3, 2);
table.get(0, 0).addParagraph().appendText("Row 1 Cell 1");
table.get(0, 1).addParagraph().appendText("Row 1 Cell 2");
table.get(1, 0).addParagraph().appendText("Row 2 Cell 1");
table.get(1, 1).addParagraph().appendText("Row 2 Cell 2");
table.get(2, 0).addParagraph().appendText("Row 3 Cell 1");
table.get(2, 1).addParagraph().appendText("Row 3 Cell 2");

// Add a new custom table style.
WTableStyle tableStyle = (WTableStyle) document.addTableStyle("CustomStyle1");
tableStyle.getTableProperties().setRowStripe((long) 1);
// Apply conditional formatting for the first row.
ConditionalFormattingStyle firstRowStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstRow);
firstRowStyle.getCharacterFormat().setBold(true);
// Apply conditional formatting for odd rows.
ConditionalFormattingStyle oddRowBandingStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.OddRowBanding);
oddRowBandingStyle.getCharacterFormat().setItalic(true);
// Apply built-in table style as base style for CustomStyle1.
tableStyle.applyBaseStyle(BuiltinTableStyle.TableContemporary);
// Apply the custom table style to the table.
table.applyStyle("CustomStyle1");
document.getLastSection().addParagraph();

// Create another table in the Word document.
table = (WTable) document.getLastSection().addTable();
table.resetCells(3, 2);
table.get(0, 0).addParagraph().appendText("Row 1 Cell 1");
table.get(0, 1).addParagraph().appendText("Row 1 Cell 2");
table.get(1, 0).addParagraph().appendText("Row 2 Cell 1");
table.get(1, 1).addParagraph().appendText("Row 2 Cell 2");
table.get(2, 0).addParagraph().appendText("Row 3 Cell 1");
table.get(2, 1).addParagraph().appendText("Row 3 Cell 2");

// Add a new custom table style.
tableStyle = (WTableStyle) document.addTableStyle("CustomStyle2");
tableStyle.getTableProperties().setRowStripe((long) 1);
// Apply conditional formatting for the first row.
firstRowStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstRow);
firstRowStyle.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
// Apply conditional formatting for odd rows.
oddRowBandingStyle = tableStyle.getConditionalFormattingStyles().add(ConditionalFormattingType.OddRowBanding);
oddRowBandingStyle.getCharacterFormat().setTextColor((ColorSupport.getRed()).clone());

// Add a new custom table style.
WTableStyle tableStyle2 = (WTableStyle) document.addTableStyle("CustomStyle3");
tableStyle2.getTableProperties().setRowStripe((long) 1);
// Apply conditional formatting for the first row.
ConditionalFormattingStyle firstRowStyle2 = tableStyle2.getConditionalFormattingStyles().add(ConditionalFormattingType.FirstRow);
firstRowStyle2.getCellProperties().setBackColor((ColorSupport.getBlue()).clone());
// Apply conditional formatting for odd rows.
ConditionalFormattingStyle oddRowStyle2 = tableStyle2.getConditionalFormattingStyles().add(ConditionalFormattingType.OddRowBanding);
oddRowStyle2.getCellProperties().setBackColor((ColorSupport.getYellow()).clone());
// Apply custom table style as base style for another custom table style.
tableStyle2.applyBaseStyle("CustomStyle2");
// Apply the custom table style to the table.
table.applyStyle("CustomStyle3");
// Save the Word document.
document.save("Sample.docx", FormatType.Docx);
// Close the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

## Merging Cells Vertically and Horizontally

You can combine two or more table cells located in the same row or column into a single cell.

The following code example illustrates how to apply horizontal merge to a specified range of cells in a specified row.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
section.addParagraph().appendText("Horizontal merging of Table cells");
IWTable table = section.addTable();
table.resetCells(5, 5);
// Specify the horizontal merge from the second cell to the fifth cell in the third row.
table.applyHorizontalMerge(2, 1, 4);
// Save and close the document instance.
document.save("HorizontalMerge.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to apply vertical merge to a specified range of rows in a specified column.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
section.addParagraph().appendText("Vertical merging of Table cells");
IWTable table = section.addTable();
table.resetCells(5, 5);
// Specify the vertical merge to the third cell, from the second row to the fifth row.
table.applyVerticalMerge(2, 1, 4);
// Save and close the document instance.
document.save("VerticalMerge.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to create a table that contains horizontally merged cells.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
section.addParagraph().appendText("Horizontal merging of Table cells");
IWTable table = section.addTable();
table.resetCells(2, 2);
// Add content to table cells.
table.get(0, 0).addParagraph().appendText("First row, First cell");
table.get(0, 1).addParagraph().appendText("First row, Second cell");
table.get(1, 0).addParagraph().appendText("Second row, First cell");
table.get(1, 1).addParagraph().appendText("Second row, Second cell");
// Specify the horizontal merge start to the first row, first cell.
table.get(0, 0).getCellFormat().setHorizontalMerge(CellMerge.Start);
// Modify the cell content.
table.get(0, 0).getParagraphs().get(0).setText("Horizontally merged cell");
// Specify the horizontal merge continuation to the second row, second cell.
table.get(0, 1).getCellFormat().setHorizontalMerge(CellMerge.Continue);
// Save and close the document instance.
document.save("HorizontalMerge.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

The following code example illustrates how to create a table with vertically merged cells.

{% tabs %}

{% highlight JAVA %}
// Create an instance of WordDocument class.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
section.addParagraph().appendText("Vertical merging of Table cells");
IWTable table = section.addTable();
table.resetCells(2, 2);
// Add content to table cells.
table.get(0, 0).addParagraph().appendText("First row, First cell");
table.get(0, 1).addParagraph().appendText("First row, Second cell");
table.get(1, 0).addParagraph().appendText("Second row, First cell");
table.get(1, 1).addParagraph().appendText("Second row, Second cell");
// Specify the vertical merge start to the first row, first cell.
table.get(0, 0).getCellFormat().setVerticalMerge(CellMerge.Start);
// Modify the cell content.
table.get(0, 0).getParagraphs().get(0).setText("Vertically merged cell");
// Specify the vertical merge continuation to the second row, first cell.
table.get(1, 0).getCellFormat().setVerticalMerge(CellMerge.Continue);
// Save and close the document instance.
document.save("VerticalMerge.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}
  
## Specifying table header row to repeat on each page

You can specify one or more rows in a table to be repeated as header rows at the top of each page when the table spans across multiple pages.

* In the case of a single header row, it must be the first row in the table.
* In the case of multiple header rows, the header rows must be consecutive from the first row of the table.

N> Heading rows do not have any effect with nested tables in Microsoft Word as well as DocIO.

The following code example illustrates how to create a table with a single header row.

{% tabs %}  

{% highlight JAVA %}
//Create an instance of WordDocument class.
WordDocument document = new WordDocument();
IWSection section = document.addSection();
IWTable table = section.addTable();
table.resetCells(50, 1);
WTableRow row = table.getRows().get(0);
//Specify the first row as a header row of the table.
row.setIsHeader(true);
row.setHeight(20);
row.setHeightType(TableRowHeightType.AtLeast);
row.getCells().get(0).addParagraph().appendText("Header Row");
for (int i = 1; i < 50; i++) {
 
    row = table.getRows().get(i);
    row.setHeight(20);
    row.setHeightType(TableRowHeightType.AtLeast);
    row.getCells().get(0).addParagraph().appendText("Text in Row" + i);
}
//Save and close the document instance.
document.save("TableWithHeaderRow.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Keeping rows from breaking across pages

You can enable or disable the table row content to split across multiple pages when the row contents do not fit in a previous page.

The following code example illustrates how to disable all the table rows from splitting across multiple pages.

{% tabs %} 

{% highlight JAVA %}
//Create an instance of WordDocument class.
WordDocument document = new WordDocument("Template.docx");
WSection section = document.getSections().get(0);
WTable table = section.getTables().get(0);
//Disable breaking across pages for all rows in the table.
for (Object row_tempObj : table.getRows()) {
 
    WTableRow row = (WTableRow) row_tempObj;
    row.getRowFormat().setIsBreakAcrossPages(false);
}
//Save and close the document instance.
document.save("Result.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

			  
  
## Iterating through table elements

The following code example illustrates how to iterate through the table and apply a background color to a particular cell.

{% tabs %} 

{% highlight JAVA %}
//Create an instance of WordDocument class.
WordDocument document = new WordDocument("Template.docx");
WSection section = document.getSections().get(0);
WTable table = section.getTables().get(0);
//Iterate the rows of the table.
for (Object row_tempObj : table.getRows()) {
 
    WTableRow row = (WTableRow) row_tempObj;
    //Iterate through the cells of rows.
    for (Object cell_tempObj : row.getCells()) {
  
        WTableCell cell = (WTableCell) cell_tempObj;
        //Iterate through the paragraphs of the cell.
        for (Object paragraph_tempObj : cell.getParagraphs()) {
   
            WParagraph paragraph = (WParagraph) paragraph_tempObj;
            //When the paragraph contains text "Panda" then apply green as the background color to the cell.
            if (paragraph.getText().contains("Panda"))
                cell.getCellFormat().setBackColor(ColorSupport.getGreen());
        }
    }
}
//Save and close the document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
				 
{% endhighlight %}

{% endtabs %}

## Removing the table

You can remove a table from a text body by its instance or by its index position in the text body item collection. The following code example shows how to remove a table in a Word document.

{% tabs %} 

{% highlight JAVA %}
//Create an instance of WordDocument class.
WordDocument document = new WordDocument("Template.docx");
//Access the instance of the first section in the Word document.
WSection section = document.getSections().get(0);
//Access the instance of the first table in the section.
WTable table = section.getTables().get(0);
//Remove a table from the text body.
section.getBody().getChildEntities().remove(table);
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}