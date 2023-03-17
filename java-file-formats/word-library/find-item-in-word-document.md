---
title: Find item in Word document in Java | Syncfusion
description: Find an item in the Word document in Java using Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---

# Find item in Word document

Just like you can search for a text in a Word document, you can also search for an item (like image, content control, textbox, and so on). The Java Word library supports to find an item in Word document based on its properties. With this functionality, you can:

* Find the first item based on one property.
* Find the first item based on multiple properties
* Find all the items based on one property.
* Find all the items based on multiple properties.

## Find item by property

Using findItemByProperty API, find the first item in the Word document that has the specified property name and value.

The following code example illustrates how to find the first item based on one property.

{% tabs %}

{% highlight JAVA %}
//Load the input Word document.
WordDocument document = new WordDocument("Input.docx",FormatType.Docx);
//Find picture by alternative text.
WPicture picture = (WPicture) document.findItemByProperty(EntityType.Picture, "AlternativeText", "Logo");
//Resize the picture.
if(picture != null)
{
    picture.setHeight((float)75);
    picture.setWidth((float)100);
}
//Save the Word document.
document.save("Sample.docx",FormatType.Docx);
//Close the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

## Find item by properties

Using findItemByProperties API, find the first item in the Word document based on multiple property names and their corresponding values.

The following code example illustrates how to find the first item in Word document based on multiple property names and their corresponding values.

{% tabs %}

{% highlight JAVA %}
//Load the input Word document.
WordDocument document = new WordDocument("Input.docx",FormatType.Docx);
String[] propertyNames=new String[]{"Title","Rows.Count"};
String[] propertyValues =new String[]{"SupplierDetails","6"};
//Find the table by Title and Rows Count.
WTable table = (WTable)document.findItemByProperties(EntityType.Table,propertyNames,propertyValues);
//Remove the table in document.
if(table != null)
   table.getOwnerTextBody().getChildEntities().remove(table);
//Save the Word document.
document.save("Sample.docx",FormatType.Docx);
//Close the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

## Find all items by property

Using findAllItemsByProperty API, find all the items in the Word document that has the specified property name and value.

The following code example illustrates how to find all the items in Word document based on one property.

{% tabs %}

{% highlight JAVA %}
//Load the input Word document.
WordDocument document = new WordDocument("Input.docx",FormatType.Docx);
//Find all footnote and endnote by EntityType in Word document.
ListSupport<Entity> footNotes = document.findAllItemsByProperty(EntityType.Footnote,null,null);
//Remove the footnotes and endnotes.
for(int i = 0;i<footNotes.getCount();i++)
{
    WFootnote footnote = (WFootnote)footNotes.get(i);
    footnote.getOwnerParagraph().getChildEntities().remove(footnote);
}
//Find all fields by FieldType.
ListSupport<Entity> fields = document.findAllItemsByProperty(EntityType.Field,"FieldType",FieldType.FieldHyperlink.toString());
//Iterate the hyperlink field and change URL.
for(int i = 0;i<fields.getCount();i++)
{
    //Create hyperlink instance from field to manipulate the hyperlink.
    Hyperlink hyperlink = new Hyperlink((WField)fields.get(i));
    //Modifies the Uri of the hyperlink.
    if(hyperlink.getType().getEnumValue()==HyperlinkType.WebLink.getEnumValue() && hyperlink.getTextToDisplay().equals("HTML"))
       hyperlink.setUri("http://www.w3schools.com/");
}
//Save the Word document.
document.save("Sample.docx",FormatType.Docx);
//Close the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

## Find all items by properties

Using findAllItemsByProperties API, find all the items in the Word document based on multiple property names and their corresponding values.

The following code example illustrates how to find all the items in Word document based on multiple property names and their corresponding values.

{% tabs %}

{% highlight JAVA %}
//Load the input Word document.
WordDocument document = new WordDocument("Input.docx",FormatType.Docx);
String[] propertyNames = {"ContentControlProperties.Title","ContentControlProperties.Tag"};
String[] propertyValues = {"CompanyName","CompanyName"};
//Find all block content controls by Title and Tag.
ListSupport<Entity> blockContentControls = document.findAllItemsByProperties(EntityType.BlockContentControl,propertyNames,propertyValues);
//Iterate the block content controls and remove the block content controls.
for(int i = 0;i<blockContentControls.getCount();i++)
{
    BlockContentControl blockContentControl = (BlockContentControl)blockContentControls.get(i);
    blockContentControl.getOwnerTextBody().getChildEntities().remove(blockContentControl);
}
propertyNames=new String[]{"ContentControlProperties.Title","ContentControlProperties.Tag"};
propertyValues=new String[]{"Contact","Contact"};
//Find all the inline content controls by Title and Tag. 
ListSupport<Entity> inlineContentControls = document.findAllItemsByProperties(EntityType.InlineContentControl,propertyNames,propertyValues);
//Iterate the inline content controls and remove the inline content controls.
for(int i = 0;i<inlineContentControls.getCount();i++)
{
    InlineContentControl inlineContentControl = (InlineContentControl)inlineContentControls.get(i);
    inlineContentControl.getOwnerParagraph().getChildEntities().remove(inlineContentControl);
}
propertyNames=new String[]{"CharacterFormat.Bold","CharacterFormat.Italic"};
propertyValues=new String[]{String.valueOf(true),String.valueOf(true)};
//Find all the bold and italic text.
ListSupport<Entity> textRanges = document.findAllItemsByProperties(EntityType.TextRange,propertyNames,propertyValues);
//Iterate the textRanges and remove the bold and italic.
for(int i = 0;i<textRanges.getCount();i++)
{
    WTextRange textRange = (WTextRange)textRanges.get(i);
    textRange.getCharacterFormat().setBold(false);
    textRange.getCharacterFormat().setItalic(false);
}
//Save the Word document.
document.save("Sample.docx",FormatType.Docx);
//Close the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

T> Passing null for both the property names and property values in the above APIs, you can also find an item in a Word document without relying on any property.