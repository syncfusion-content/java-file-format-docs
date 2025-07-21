---
title: Accepting or Rejecting Track Changes | Word Library | Syncfusion
description: This section illustrates how to accept or reject the track changes in the Word document using the Syncfusion Word library (Essential DocIO).
platform: java-file-formats
control: Word Library
documentation: UG
---
# Accepting or Rejecting Track Changes

It is used to keep track of the changes made to a Word document. It helps to maintain the record of the author, name, and time for every insertion, deletion, or modification in a document. This can be enabled by using the `TrackChanges` property of the Word document.

N> 
With this support, the changes made in the Word document by the DocIO library cannot be tracked.

The following code example illustrates how to enable track changes in the document.

{% tabs %}   

{% highlight JAVA %}
// Creates a new Word document.
WordDocument document = new WordDocument();
// Adds a new section to the document.
IWSection section = document.addSection();
// Adds a new paragraph to the section.
IWParagraph paragraph = section.addParagraph();
// Appends text to the paragraph.
IWTextRange text = paragraph.appendText("This sample illustrates how to track the changes made to the Word document. ");
// Sets font name and size for the text.
text.getCharacterFormat().setFontName("Times New Roman");
text.getCharacterFormat().setFontSize((float)14);
text = paragraph.appendText("This track changes is useful in a shared environment.");
text.getCharacterFormat().setFontSize((float)12);
// Turns on the track changes option.
document.setTrackChanges(true);
// Saves and closes the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %} 

## Accept all changes

You can **accept all track changes in a Word document** using the `acceptAll` method.

The following code example shows how to accept all the tracked changes.

{% tabs %}   

{% highlight JAVA %}
// Opens an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Accepts all the tracked changes revisions.
if (document.getHasChanges())
     document.getRevisions().acceptAll();
// Saves and closes the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}  

{% endtabs %}

By executing the above code example, it generates an output Word document as follows.

![Accepting all track changes in Word document](WorkingWithTrackChanges_images/AcceptAll.png)

## Reject all changes

You can **reject all track changes in a Word document** using the `rejectAll` method.

The following code example shows how to reject all the tracked changes.

{% tabs %}   

{% highlight JAVA %}
// Opens an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Rejects all the tracked changes revisions.
if (document.getHasChanges())
    document.getRevisions().rejectAll();
// Saves and closes the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

By executing the above code example, it generates an output Word document as follows.

![Rejecting all track changes in Word document](WorkingWithTrackChanges_images/RejectAll.png)

## Accept all changes by a particular reviewer

You can **accept all changes made by the author** in the Word document using the `accept` method.

The following code example shows how to accept the tracked changes made by the author.

{% tabs %}   

{% highlight JAVA %}
// Opens an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Iterates into all the revisions in the Word document.
for (int i = document.getRevisions().getCount() - 1; i >= 0; i--) 
{
    // Checks the author of the current revision and accepts it.
    if (document.getRevisions().get(i).getAuthor().equals("Nancy Davolio"))
        document.getRevisions().get(i).accept();
    // Resets to the last item when accepting the moving related revisions.
    if (i > document.getRevisions().getCount() - 1)
        i = document.getRevisions().getCount();
}
// Saves and closes the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

## Reject all changes by a particular reviewer

You can **reject all changes made by the author** in the Word document using the `reject` method.

The following code example shows how to reject the tracked changes made by the author.

{% tabs %}   

{% highlight JAVA %}
// Opens an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Iterates into all the revisions in the Word document.
for (int i = document.getRevisions().getCount() - 1; i >= 0; i--) 
{
    // Checks the author of the current revision and rejects it.
    if (document.getRevisions().get(i).getAuthor().equals("Nancy Davolio"))
        document.getRevisions().get(i).reject();
    // Resets to the last item when rejecting the moving related revisions.
    if (i > document.getRevisions().getCount() - 1)
        i = document.getRevisions().getCount();
}
// Saves and closes the document.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %} 

{% endtabs %}

## Revision information

You can get the **revision information of track changes** in the Word document like the author name, date, and type of revision.

The following code example shows how to get the details about the revision information of track changes.

{% tabs %}   

{% highlight JAVA %}
// Opens an existing Word document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
// Accesses the first revision in the Word document.
Revision revision = document.getRevisions().get(0);
// Gets the name of the user who made the specified tracked change.
String author = revision.getAuthor();
// Gets the date and time that the tracked change was made.
LocalDateTime dateTime = revision.getDate();
// Gets the type of the track changes revision.
RevisionType revisionType = revision.getRevisionType();
// Closes the document.
document.close();
{% endhighlight %} 

{% endtabs %}

Frequently Asked Questions

* [How to check whether a Word document contains tracked changes or not?](https://help.syncfusion.com/document-processing/word/word-library/java/faq#how-to-check-whether-a-word-document-contains-tracked-changes-or-not)
* [How to accept or reject track changes of specific type in the Word document?](https://help.syncfusion.com/document-processing/word/word-library/java/faq#how-to-accept-or-reject-track-changes-of-specific-type-in-the-word-document)
* [How to enable track changes for Word document?](https://help.syncfusion.com/document-processing/word/word-library/java/faq#how-to-enable-track-changes-for-word-document)