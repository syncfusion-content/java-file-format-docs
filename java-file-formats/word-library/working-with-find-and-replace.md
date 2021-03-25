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