---
title: Applying Watermark | Syncfusion
description: This section illustrates how to insert text or pictures watermark to the Word document using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Working with Watermark

Watermarks are text or pictures that appear behind the document text. You can access the watermark in the document by using the `Watermark` property of `WordDocument` class.

There are two types of watermarks: Text and Picture.

## Text watermark

You can add or modify text watermark in the Word document. The `TextWatermark` class represents the text watermark in the Word document.

The following code example shows how to add a text watermark to the Word document.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add a section and a paragraph in the document.
document.ensureMinimal();
IWParagraph paragraph = document.getLastParagraph();
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Create a new text watermark.
TextWatermark textWatermark = new TextWatermark("TextWatermark", "", 250, 100);
//Set the created watermark to the document.
document.setWatermark(textWatermark);
//Set the text watermark font size.
textWatermark.setSize(72);
//Set the text watermark layout to horizontal.
textWatermark.setLayout(WatermarkLayout.Horizontal);
textWatermark.setSemitransparent(false);
//Set the text watermark text color.
textWatermark.setColor(ColorSupport.getBlack());
//Save the Word document.
document.save("Result_watermark1.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Picture Watermark

You can add or modify the picture watermark in the Word document. The `PictureWatermark` class represents the picture watermark in the Word document.

The following code example shows how to add a picture watermark to the Word document.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document.
WordDocument document = new WordDocument();
//Add a section and a paragraph in the document.
document.ensureMinimal();
IWParagraph paragraph = document.getLastParagraph();
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Create a new picture watermark.
PictureWatermark picWatermark = new PictureWatermark();
//Set the scaling to picture.
picWatermark.setScaling(120f);
picWatermark.setWashout(true);
//Set the picture watermark to document.
document.setWatermark(picWatermark);
//Set the image to the picture watermark.
Path path = Paths.get("David.png");
byte[] data = Files.readAllBytes(path);
picWatermark.loadPicture(data);
//Save and close the document.
document.save("PictureWatermark.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}