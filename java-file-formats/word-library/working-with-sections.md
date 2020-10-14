---
title: Working with Sections | Syncfusion
description: This section illustrates how to Work with Sections in Word document using Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Sections

A section contains the contents present in the headers, footers, and the main document using the instances of `WTextBody`. A section also has a specific set of properties used to define the page settings, a number of columns, headers, and footers, and more that decide how the text appears. The `WTextBody` represents a group of paragraphs, tables, and more.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

You can add the multiple sections to the document. When you add more than one section into the word document, the section starts from the next page by default.

You can also add a new section that starts on the same page by specifying the `BreakCode` as shown in the following code example.

{% tabs %}   

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a paragraph to the created section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append the text to the created paragraph.
paragraph.appendText(paraText);
//Add the new section to the document.
section = document.addSection();
//Set a section break.
section.setBreakCode(SectionBreakCode.NoBreak) ;
//Add a paragraph to the created section.
paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText(paraText); 
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %} 

## Specifying Page Properties

Each section has its page setup properties such as page size, orientation, margins, borders, and more.

The following code example shows how to set the page setup properties.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add the section into the Word document.
IWSection section = document.addSection();
//Set the page setup options.
section.getPageSetup().setOrientation(PageOrientation.Landscape);
section.getPageSetup().getMargins().setAll(72);
section.getPageSetup().getBorders().setLineWidth(2);
//Add a paragraph to the created section.
IWParagraph paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company."); 
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  
 
## Creating Multi-column document

You can split the contents into two or more columns by specifying the column width and spacing between columns.

The following code example shows how to display the contents in multiple columns.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add the section into the Word document.
IWSection section = document.addSection();
//Add the column into the section.
section.addColumn(150, 20);
//Add the column into the section.
section.addColumn(150, 20);
//Add the column into the section.
section.addColumn(150, 20);
//Add a paragraph to the created section.
IWParagraph paragraph = section.addParagraph();
//Add a paragraph to the created section.
paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append the text to the created paragraph.
paragraph.appendText(paraText);
//Add the column break.
paragraph.appendBreak(BreakType.ColumnBreak);
//Add a paragraph to the created section.
paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText(paraText);
//Add the column break.
paragraph.appendBreak(BreakType.ColumnBreak);
//Add a paragraph to the created section.
paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText(paraText);
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %} 

## Creating document with different page settings

You can prefer to have more sections in a Word document when you need to have different page settings or headers and footers for a specific set of contents. The following code example shows how to create a Word document with the multiple sections whose page orientation is portrait and landscape respectively.

{% tabs %} 

{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Add the section into the Word document.
IWSection section = document.addSection();
//Add a paragraph to the created section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append the text to the created paragraph.
paragraph.appendText(paraText);
//Set the page orientation as a portrait.
section.getPageSetup().setOrientation(PageOrientation.Portrait);
//Add the new section to the document.
section = document.addSection();
//Set the section break.
section.setBreakCode(SectionBreakCode.NewPage) ;
paragraph = section.addParagraph();
//Set the page orientation as a landscape
section.getPageSetup().setOrientation(PageOrientation.Landscape);
//Append the text to the paragraph.
paragraph.appendText(paraText);
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}
   
## Working with Headers and Footers

The header and footer also represent the group of paragraphs and tables that occur at the top and bottom of the page respectively. The header and footer may vary for each section. The following are the types of Headers or Footers:

* FirstPageHeader: Represents the first-page header of the document.
* FirstPageFooter: Represents the first-page footer of the document.
* OddHeader: Represents the odd page header of the document and it is the default header for the section.
* OddFooter: Represents the odd page footer of the document and it is the default footer for the section.
* EvenHeader: Represents the even page header of the document.
* Even Footer: Represents the even page footer of the document.

The following code example illustrates how to add simple header and footer into a Word document.

{% tabs %} 

{% highlight JAVA %}
//Create a new document.
WordDocument document = new WordDocument();
//Add the first section to the document.
IWSection section = document.addSection();
//Add a paragraph to the section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append some text to the first page of the document.
paragraph.appendText("\r\r[ First Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the second page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Second Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the third page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Third Page ] \r\r" + paraText);
//Insert the default page header.
paragraph = section.getHeadersFooters().getOddHeader().addParagraph();
paragraph.appendText("[ Default Page Header ]");
//Insert the default page footer.
paragraph = section.getHeadersFooters().getOddHeader().addParagraph();
paragraph.appendText("[ Default Page Footer ]");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

You can have a specific header and footer contents for the first page in a Word document. The following code illustrates the same.

{% tabs %} 

{% highlight JAVA %}
//Create a new document.
WordDocument document = new WordDocument();
//Add the first section to the document.
IWSection section = document.addSection();
//Set the DifferentFirstPage as a true for inserting the header and footer text.
section.getPageSetup().setDifferentFirstPage(true);
//Add a paragraph to the section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append some text to the first page of the document.
paragraph.appendText("\r\r[ First Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the second page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Second Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the third page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Third Page ] \r\r" + paraText);
//Insert the first page header.
paragraph = section.getHeadersFooters().getFirstPageHeader().addParagraph();
paragraph.appendText("[First Page Header ]");
//Insert the first page footer.
paragraph = section.getHeadersFooters().getFirstPageFooter().addParagraph();
paragraph.appendText("[ First Page Footer ]");
//Insert the default page header.
paragraph = section.getHeadersFooters().getOddHeader().addParagraph();
paragraph.appendText("[ Default Page Header ]");
//Insert the default page footer.
paragraph = section.getHeadersFooters().getOddFooter().addParagraph();
paragraph.appendText("[ Default Page Footer ]");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

A word document can have a different header and footer for the odd and even pages.

The following code example shows how to set a different header and footer for the odd and even pages of the document.

{% tabs %} 

{% highlight JAVA %}
//Create a new document.
WordDocument document = new WordDocument();
//Add the first section to the document.
IWSection section = document.addSection();
//Set the DifferentOddAndEvenPages to true for inserting the header and footer text.
section.getPageSetup().setDifferentOddAndEvenPages(true);
//Add a paragraph to the section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append some text to the first page of the document.
paragraph.appendText("\r\r[ First Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the second page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Second Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the third page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Third Page ] \r\r" + paraText);
//Insert the odd page header.
paragraph = section.getHeadersFooters().getOddHeader().addParagraph();
paragraph.appendText("[ Odd Page Header ]");
//Insert the default page footer.
paragraph = section.getHeadersFooters().getOddFooter().addParagraph();
paragraph.appendText("[ Odd Page Footer ]");
//Insert the even page header.
paragraph = section.getHeadersFooters().getEvenHeader().addParagraph();
paragraph.appendText("[Even Page Header ]");
//Insert the even page footer.
paragraph = section.getHeadersFooters().getEvenFooter().addParagraph();
paragraph.appendText("[ Even Page Footer ]");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

You can use the previous section header and footer for the current section by using the `LinkToPrevious` property.

The following code example shows how to link the previous section header and footer for the current section.

{% tabs %}  

{% highlight JAVA %}
//Create a new document.
WordDocument document = new WordDocument();
//Add the first section to the document.
IWSection section = document.addSection();
//Insert the first section header.
section.getHeadersFooters().getHeader().addParagraph().appendText("[ First Section Header ]");
//Insert the first section footer.
section.getHeadersFooters().getFooter().addParagraph().appendText("[ First Section Footer ]");
//Add a paragraph to the section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append some text to the first page of the document.
paragraph.appendText("\r\r[ First Page ] \r\r" + paraText);
//Add the second section to the document.
section = document.addSection();
//Insert the second section header.
section.getHeadersFooters().getHeader().addParagraph().appendText("[ Second Section Header ]");
//Insert the second section footer.
section.getHeadersFooters().getFooter().addParagraph().appendText("[ Second Section Footer ]");
//Set the LinkToPrevious to true for retrieve the header and footer from the previous section.
section.getHeadersFooters().setLinkToPrevious(true);
//Append some text to the second page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Second Page ] \r\r" + paraText);
//Add the third section to the document.
section = document.addSection();
//Insert the third section header.
section.getHeadersFooters().getHeader().addParagraph().appendText("[ Third Section Header ]");
//Insert the third section footer.
section.getHeadersFooters().getFooter().addParagraph().appendText("[ Third Section Footer ]");
//Append some text to the third page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Third Page ] \r\r" + paraText);
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Adding Page Numbers

You can insert the current page number within the document contents. The following code example shows how to insert the current page number within the footer.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add the section into a Word document.
IWSection section = document.addSection();
section.getPageSetup().setPageStartingNumber(1);
section.getPageSetup().setRestartPageNumbering(true);
section.getPageSetup().setPageNumberStyle(PageNumberStyle.Arabic);
//Add a footer paragraph text to the document.
IWParagraph paragraph = section.getHeadersFooters().getFooter().addParagraph();
paragraph.getParagraphFormat().getTabs().addTab(523f, TabJustification.Right, TabLeader.NoLeader);
//Add text for the footer paragraph.
paragraph.appendText("Copyright Northwind Inc. 2001 - 2015");
//Add the page number field to the document.
paragraph.appendText("\tPage ");
paragraph.appendField("Page", FieldType.FieldPage);
//Add the paragraph.
paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

The following code example shows how to add the current page number and the total number of pages in the header or footer.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add the section into the Word document.
IWSection section = document.addSection();
section.getPageSetup().setPageStartingNumber(1);
section.getPageSetup().setRestartPageNumbering(true);
section.getPageSetup().setPageNumberStyle(PageNumberStyle.Arabic);
//Add a footer paragraph text to the document.
IWParagraph paragraph = section.getHeadersFooters().getFooter().addParagraph();
paragraph.getParagraphFormat().getTabs().addTab(523f, TabJustification.Right, TabLeader.NoLeader);
// Add the text for the footer paragraph.
paragraph.appendText("Copyright Northwind Inc. 2001 - 2015\t");
//Add the text.
paragraph.appendText(" Page ");
//Add the page number field to the document.
paragraph.appendField("CurrentPageNumber", FieldType.FieldPage);
// Add the text.
paragraph.appendText(" of ");
//Add the number of pages field to the document.
paragraph.appendField("TotalNumberOfPages", FieldType.FieldNumPages);
//Add the paragraph.
paragraph = section.addParagraph();
//Append the text to the created paragraph.
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

The following code example shows how to adjust the height of the header and footer.

{% tabs %} 

{% highlight JAVA %}
//Create a new document.
WordDocument document = new WordDocument();
//Add the first section to the document.
IWSection section = document.addSection();
//Specify the value to the header distance.
section.getPageSetup().setHeaderDistance(100);
//Specify the value to the footer distance.
section.getPageSetup().setFooterDistance(100);
//Add a paragraph to the section.
IWParagraph paragraph = section.addParagraph();
String paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
//Append some text to the first page of the document.
paragraph.appendText("\r\r[ First Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the second page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Second Page ] \r\r" + paraText);
paragraph.getParagraphFormat().setPageBreakAfter(true);
//Append some text to the third page of the document.
paragraph = section.addParagraph();
paragraph.appendText("\r\r[ Third Page ] \r\r" + paraText);
//Insert the default page header.
paragraph = section.getHeadersFooters().getOddHeader().addParagraph();
paragraph.appendText("[ Default Page Header ]");
//Insert the default page footer.
paragraph = section.getHeadersFooters().getOddFooter().addParagraph();
paragraph.appendText("[ Default Page Footer ]");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Removing a Section

The following code example shows how to remove a particular section from the word document.

{% tabs %}  

{% highlight JAVA %}
//Open an input Word template.
WordDocument document = new WordDocument("inputFileName");
//Remove the second section from the collection.
document.getSections().removeAt(1);
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}