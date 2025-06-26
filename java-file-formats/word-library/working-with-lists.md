---
title: Working with Lists | Syncfusion
description: This section explains how to work with the lists in Word document using Syncfusion Java Word library (Essential DocIO) 
platform: java-file-formats
control: Word Library
documentation: UG
---

# Working with lists

Lists can organize and format the contents of a document in hierarchical way. There are nine levels in the list, starting from level 0 to level 8. DocIO supports both built-in list styles and custom list styles. The following are the types of list supported in DocIO: 

* Numbered list
* Bulleted list 

## Create Bulleted List

The following code example explains how to create a simple bulleted list.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Apply the default numbered list style.
paragraph.getListFormat().applyDefBulletStyle();
//Add text to the paragraph.
paragraph.appendText("List item 1");
//Continue the list defined.
paragraph.getListFormat().continueListNumbering();
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 2");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 3");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Save the Word document.
document.save("simple bulleted list.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Create Numbered List

The following code example explains how to create a simple numbered list.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Apply the default numbered list style.
paragraph.getListFormat().applyDefNumberedStyle();
//Add the text to the paragraph.
paragraph.appendText("List item 1");
//Continue the list defined.
paragraph.getListFormat().continueListNumbering();
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 2");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 3");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Create Multilevel List

The following code example explains how to create a multilevel bulleted list.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Apply the default numbered list style.
paragraph.getListFormat().applyDefBulletStyle();
//Add the text to the paragraph.
paragraph.appendText("List item 1 - Level 0");
//Continue the list defined.
paragraph.getListFormat().continueListNumbering();
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 2 - Level 1");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 3 - Level 2");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Create Multilevel Numbered List

The following code example explains how to create multilevel numbered list.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Apply the default numbered list style.
paragraph.getListFormat().applyDefNumberedStyle();
//Add the text to the paragraph.
paragraph.appendText("List item 1 - Level 0");
//Continue the list defined.
paragraph.getListFormat().continueListNumbering();
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 2 - Level 1");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("List item 3 - Level 2");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## List number format

The ListPatternType enum in DocIO lets you customize how list numbers appear in Word documents. It supports 61 styles, including Arabic, Hebrew, and more. This is useful for creating region-specific documents or applying culturally appropriate numbering formats.

The following code example demonstrates how to create a list number format.

{% tabs %} 

{% highlight JAVA %}
// Creates a new Word document.
WordDocument document = new WordDocument();
// Adds a new section to the document.
IWSection section = document.addSection();

// Adds a numbered list style with CardinalText pattern (One, Two, Three, ...).
ListStyle listStyle = document.addListStyle(ListType.Numbered, "CardinalText");
WListLevel levelOne = listStyle.getLevels().get(0);
levelOne.setPatternType(ListPatternType.CardinalText);
levelOne.setStartAt(1);
// Adds a heading paragraph for the CardinalText list.
IWParagraph paragraph = section.addParagraph();
paragraph.appendText("List pattern Cardinal Text");
// Adds first list item using CardinalText style.
paragraph = section.addParagraph();
paragraph.appendText("List item 1");
paragraph.getListFormat().applyStyle("CardinalText");
paragraph.getListFormat().continueListNumbering();
// Adds second list item using CardinalText style.
paragraph = section.addParagraph();
paragraph.appendText("List item 2");
paragraph.getListFormat().applyStyle("CardinalText");
paragraph.getListFormat().continueListNumbering();
// Adds third list item using CardinalText style.
paragraph = section.addParagraph();
paragraph.appendText("List item 3");
paragraph.getListFormat().applyStyle("CardinalText");
paragraph.getListFormat().continueListNumbering();
// Adds a blank paragraph before the next list.
section.addParagraph();

// Adds a numbered list style with HindiLetter1 pattern.
listStyle = document.addListStyle(ListType.Numbered, "HindiLetter1");
levelOne = listStyle.getLevels().get(0);
levelOne.setPatternType(ListPatternType.HindiLetter1);
levelOne.setStartAt(1);
// Adds a heading paragraph for the HindiLetter1 list.
paragraph = section.addParagraph();
paragraph.appendText("List pattern Hindi Letter");
// Adds first list item using HindiLetter1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 1");
paragraph.getListFormat().applyStyle("HindiLetter1");
paragraph.getListFormat().continueListNumbering();
// Adds second list item using HindiLetter1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 2");
paragraph.getListFormat().applyStyle("HindiLetter1");
paragraph.getListFormat().continueListNumbering();
// Adds third list item using HindiLetter1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 3");
paragraph.getListFormat().applyStyle("HindiLetter1");
paragraph.getListFormat().continueListNumbering();
// Adds a blank paragraph before the next list.
section.addParagraph();

// Adds a numbered list style with Hebrew1 pattern.
listStyle = document.addListStyle(ListType.Numbered, "Hebrew1");
levelOne = listStyle.getLevels().get(0);
levelOne.setPatternType(ListPatternType.Hebrew1);
levelOne.setStartAt(1);
// Adds a heading paragraph for the Hebrew1 list.
paragraph = section.addParagraph();
paragraph.appendText("List pattern Hebrew");
// Adds first list item using Hebrew1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 1");
paragraph.getListFormat().applyStyle("Hebrew1");
paragraph.getListFormat().continueListNumbering();
// Adds second list item using Hebrew1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 2");
paragraph.getListFormat().applyStyle("Hebrew1");
paragraph.getListFormat().continueListNumbering();
// Adds third list item using Hebrew1 style.
paragraph = section.addParagraph();
paragraph.appendText("List item 3");
paragraph.getListFormat().applyStyle("Hebrew1");
paragraph.getListFormat().continueListNumbering();

// Saves the Word document
document.save("Sample.docx", FormatType.Docx);
// Closes the document
document.close();
{% endhighlight %}

{% endtabs %}  


N> Except for the following [ListPatternType](https://help.syncfusion.com/cr/document-processing/Syncfusion.DocIO.DLS.ListPatternType.html) enumeration values: Arabic, Bullet, ChineseCountingThousand, FarEast, KanjiDigit, LeadingZero, LowLetter, LowRoman, None, Number, Ordinal, OrdinalText, Special, UpLetter, and UpRoman, all other [ListPatternType](https://help.syncfusion.com/cr/document-processing/Syncfusion.DocIO.DLS.ListPatternType.html) values are supported only during DOCX to DOCX conversions.

## Change List Levels

The list levels can be incremented or decremented by using the `increaseIndentLevel` and `decreaseIndentLevel` methods respectively. The following code example explains how to increase or decrease the list indent levels.

{% tabs %} 

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection();
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
//Apply the default numbered list style.
paragraph.getListFormat().applyDefNumberedStyle();
//Add the text to the paragraph.
paragraph.appendText("Multilevel numbered list - Level 0");
//Continue the list defined.
paragraph.getListFormat().continueListNumbering();
//Add the second paragraph
paragraph = section.addParagraph();
paragraph.appendText("Multilevel numbered list - Level 1");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("Multilevel numbered list - Level 0");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().decreaseIndentLevel();
//Add a new paragraph.
paragraph = section.addParagraph();
paragraph.appendText("Multilevel numbered list - Level 1");
//Continue he last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Customize List

The following code example explains how to create user defined list styles.

{% tabs %}  

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection(); 
//Add a new list style to the document.          
ListStyle listStyle = document.addListStyle(ListType.Numbered, "UserDefinedList");
WListLevel levelOne = listStyle.getLevels().get(0);
//Define the follow character, prefix, suffix, and start index for level 0.
levelOne.setFollowCharacter(FollowCharacterType.Tab);
levelOne.setNumberPrefix("(");
levelOne.setNumberSufix(")");
levelOne.setPatternType(ListPatternType.LowRoman);
levelOne.setStartAt(1);
levelOne.setTabSpaceAfter(5);
levelOne.setNumberAlignment(ListNumberAlignment.Center);
WListLevel levelTwo = listStyle.getLevels().get(1);
//Define the follow character, suffix, pattern, and start index for level 1.
levelTwo.setFollowCharacter(FollowCharacterType.Tab);
levelTwo.setNumberPrefix("}");
levelTwo.setPatternType(ListPatternType.LowLetter);
levelTwo.setStartAt(2); 
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();  
//Add the text to the paragraph.
paragraph.appendText("User defined list - Level 0");
//Apply the default numbered list style.
paragraph.getListFormat().applyStyle("UserDefinedList");
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("User defined list - Level 1");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Numbered List with Prefix

The following code example explains how to create numbered list with prefix from previous level.

N> The `NumberPrefix` value for the numbered list should meet the syntax “\u000N” to update the previous list level value as a prefix to the current list level. For example, it should be represented as (“\u0000.” or “\u0000.\u0001.”).
{% tabs %}  

{% highlight JAVA %}
//Create a new Word document. 
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.addSection(); 
//Add a new list style to the document.          
ListStyle listStyle = document.addListStyle(ListType.Numbered, "UserDefinedList");
WListLevel levelOne = listStyle.getLevels().get(0);
//Define the follow character, prefix from previous level, and start index for level 0.
levelOne.setFollowCharacter(FollowCharacterType.Nothing);
levelOne.setPatternType(ListPatternType.Arabic);
levelOne.setStartAt(1);
WListLevel levelTwo = listStyle.getLevels().get(1);
//Define the follow character, prefix from previous level, pattern, and start index for level 1.
levelTwo.setFollowCharacter(FollowCharacterType.Nothing);
levelTwo.setNumberPrefix("\u0000.");
levelTwo.setPatternType(ListPatternType.Arabic);
levelTwo.setStartAt(1);
WListLevel levelThree = listStyle.getLevels().get(2);
//Define the follow character, prefix from previous level, pattern, and start index for level 1.
levelThree.setFollowCharacter(FollowCharacterType.Nothing);
levelThree.setNumberPrefix("\u0000.\u0001.");
levelThree.setPatternType(ListPatternType.Arabic);
levelThree.setStartAt(1); 
//Add a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();  
//Add a text to the paragraph.
paragraph.appendText("User defined list - Level 0");
//Apply the default numbered list style.
paragraph.getListFormat().applyStyle("UserDefinedList");
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("User defined list - Level 1");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Add the second paragraph.
paragraph = section.addParagraph();
paragraph.appendText("User defined list - Level 2");
//Continue the last defined list.
paragraph.getListFormat().continueListNumbering();
//Increase the level indent.
paragraph.getListFormat().increaseIndentLevel();
//Save the Word document.
document.save("Sample.docx", FormatType.Docx);
//Close the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Get list value

You can get the string that represents the appearance of **list value of the paragraph** in the Word document using the `ListString` API. 

This API holds the static string of the list value recently calculated while saving the document as Text. It is not updated automatically for each modification done in the Word document. Hence, you should either invoke the `getText()` method of `WordDocument` or save the Word document as Text to get the actual list value from this API.

The following example shows how to **get a string that represents the appearance of list value of the paragraph**.

{% tabs %}  

{% highlight JAVA %}
//Load an existing Word document.
WordDocument document = new WordDocument("Template.docx");
//Get the document text.
document.getText();
//Get the string that represents the appearance of list value of the paragraph.
String listString = document.getLastParagraph().getListString();
//Save and close the WordDocument instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

N> For a picture bulleted list, the `ListString` API is not valid and it will return an empty string.
