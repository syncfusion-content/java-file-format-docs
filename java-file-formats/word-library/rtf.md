---
title: RTF conversions | Word library | Syncfusion
description: This section illustrates how to perform RTF to Word conversion and Word to RTF conversions using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Word to RTF and RTF to Word Conversions

## RTF
The [Rich Text Format (RTF)](http://en.wikipedia.org/wiki/Rich_Text_Format#) is one of the document formats supported by Microsoft Word and many other Word processing applications. RTF is human readable file format invented for interchanging formatted text between applications. It is an optional format for Word that retains most formatting and all content of the original document.

Most of the Word processors (including Microsoft Word) uses the XML-based file formats, Microsoft has discontinued enhancements to the RTF and still supporting it. The last version was 1.9.1 released in 2008.

The Essential DocIO converts the RTF document into Word document and vice versa. The following code shows how to convert RTF document into Word document.

{% tabs %}
{% highlight JAVA %}
//Load an existing document.
WordDocument document = new WordDocument("Input.rtf", FormatType.Rtf);
//Save the Word document as RTF file.
document.save("RtfToWord.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

The following code example shows how to convert Word document into RTF document.

{% tabs %}
{% highlight JAVA %}
//Load an existing document.
WordDocument document = new WordDocument("Input.docx", FormatType.Docx);
//Save the Word document as RTF file.
document.save("WordToRtf.rtf", FormatType.Rtf);
//Close the document.
document.close();
{% endhighlight %}
{% endtabs %}

## Supported and Unsupported features
The supported and unsupported features of DocIO based on file formats can be referred [here](https://help.syncfusion.com/java-file-formats/word-library/supported-and-unsupported-features#)