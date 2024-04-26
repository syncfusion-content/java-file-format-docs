---
title: Word document to Text Conversion | Word library | Syncfusion
description: This section illustrates how to perform Word document to Text conversion using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Word to Text and Text to Word Conversions

The Essential DocIO converts the Word document into the Text file and vice versa. The following code example shows how to convert the Word document into text file.

{% tabs %}
{% highlight JAVA %}
//Load a template document.
WordDocument document = new WordDocument("Template.docx");
//Save the document as a text file.
document.save("WordToText.txt", FormatType.Txt);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

The following code example shows how to convert a text file into a Word document.

{% tabs %}
{% highlight JAVA %}
//Load a text file.
WordDocument document = new WordDocument("Template.txt");
//Save the document as a text file.
document.save("TextToWord.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

The following code example shows how to retrieve the Word document contents as a plain text.

{% tabs %}
{% highlight JAVA %}
//Load a template document.
WordDocument document = new WordDocument("Template.docx");
//Get the document text.
String text = document.getText();
//Create a new Word document.
WordDocument newdocument = new WordDocument();
//Add a new section.
IWSection section = newdocument.addSection();
//Add a new paragraph.
IWParagraph paragraph = section.addParagraph();
//Append the text to the paragraph.
paragraph.appendText(text);
//Save and close the document.
newdocument.save("Sample.docx");
newdocument.close();
document.close();
{% endhighlight %}
{% endtabs %}
