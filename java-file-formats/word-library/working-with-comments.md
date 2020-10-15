---
title: Working with Comments | Syncfusion
description: This section illustrates how to add, modify and remove the comments using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Comments

A comment is a note or annotation that an author or reviewer can add to a document. DocIO represents comment with `WComment` instance.

N>  The comment start and end ranges and dates can be preserved only on processing an existing document that already contains these information for each comment.

## Adding a Comment

You can add a new comment to the Word document by using `appendComment` method of `WParagraph` class. 

The following code shows how to add a new comment to the document:

{% tabs %}  

{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds a section and a paragraph in the document.
document.ensureMinimal();
IWParagraph paragraph = document.getLastParagraph();
//Appends text to the paragraph.
paragraph.appendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Adds comment to a paragraph.
WComment comment = paragraph.appendComment("comment test");
//Specifies the author of the comment.
comment.getFormat().setUser("Peter");
//Specifies the initial of the author.
comment.getFormat().setUserInitials("St");
//Set the date and time for comment.
comment.getFormat().setDateTime(LocalDateTime.now());
//Saves the Word document.
document.save("Comment.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %} 

{% endtabs %}  

## Modifying a Comment

The following code illustrates how to modify the text of an existing comment in the Word document:

{% tabs %}  

{% highlight JAVA %}
//Opens the template document.
WordDocument document = new WordDocument("Comment.docx", FormatType.Docx);
//Iterates the comments in the Word document.
for (Object comments : document.getComments())
{
	WComment comment = (WComment)comments;
	//Modifies the last paragraph text of an existing comment when it is added by "Peter".
	if ((comment.getFormat().getUser()).equals("Peter"))
		comment.getTextBody().getLastParagraph().setText("Modified Comment Content");
}
//Saves the Word document.
document.save("ModifiedComment.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}  
  
## Removing Comments

You can either remove all the comments or a particular comment from the Word document.

The following code shows how to remove all the comments in Word document.

{% tabs %}  

{% highlight JAVA %}
//Opens the template document.
WordDocument document = new WordDocument("Comment.docx", FormatType.Docx);
//Removes all the comments in a Word document.
document.getComments().clear();
//Saves the Word document.
document.save("Result.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}  

The following code shows how to remove a particular comment from Word document.

{% tabs %} 

{% highlight JAVA %}
//Opens the template document.
WordDocument document = new WordDocument("Comments.docx", FormatType.Docx);
//Removes second comments from a document.
document.getComments().removeAt(1);
//Saves the Word document.
document.save("Result.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}