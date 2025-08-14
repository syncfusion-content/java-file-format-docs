---
title: Working with Security | Word Library | Syncfusion
description: This section illustrates how to encrypt, decrypt, and protect a Word document using the Syncfusion Word Library (Essential DocIO).
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Security

You can encrypt a Word document with a password to restrict unauthorized access. You can also control the types of changes you make to this document.

## Encrypting with a password

The following code example shows how to encrypt a Word document with a password.

{% tabs %}  

{% highlight JAVA %}
// Open an input Word document.
WordDocument document = new WordDocument("Template.docx");
// Encrypt the Word document with a password.
document.encryptDocument("password");
// Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Opening the encrypted Word document

The following code example shows how to open the encrypted Word document. 

{% tabs %}  

{% highlight JAVA %}
// Open an input Word document.
WordDocument document = new WordDocument("Template.docx", "password");
// Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

## Remove encryption

You can open the encrypted Word document and remove the encryption from the document. The following code example shows how to remove the encryption from an encrypted Word document.

{% tabs %}  

{% highlight JAVA %}
// Open an encrypted Word document.
WordDocument document = new WordDocument("Template.docx", "password");
// Remove encryption in the Word document.
document.removeEncryption();
// Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Protecting the Word document from editing

You can restrict a Word document from editing either by providing a password or without a password. 

The following are the types of protection:

1. `AllowOnlyComments`: You can add/modify only the comments in the Word document.

2. `AllowOnlyFormFields`: You can modify the form field values in the Word document.

3. `AllowOnlyRevisions`: You can accept or reject the revisions in the Word document.

4. `AllowOnlyReading`: You can only view the content in the Word document.

5. `NoProtection`: You can access/edit the Word document contents as usual.

The following code example shows how to restrict editing to modify only form fields in a Word document.

{% tabs %}  

{% highlight JAVA %}

// Open a Word document.
WordDocument document = new WordDocument("Template.docx");
// Set the protection with a password and allow only modification of the form fields type.
document.protect(ProtectionType.AllowOnlyFormFields, "password"); 
// Save the Word document.
document.save("Protection.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Editable ranges

An **editable range** is a portion of a Word document that allows editing even when the document is protected. In the Syncfusion .NET Word Library (DocIO), editable ranges are represented using the **EditableRange** class. You can define these ranges programmatically to allow user edits within protected documents.

### Add an editable range

You can add an editable range to a Word document using the **appendEditableRangeStart()** and **appendEditableRangeEnd()** methods of the **WParagraph** class.

The following code example illustrates how to add an editable range in a Word document.

N> DocIO supports editable ranges in DOCX format documents only.

{% tabs %}  

{% highlight JAVA %}
// Create a Word document
WordDocument document = new WordDocument();
// Ensure at least one section and one paragraph exists
document.ensureMinimal();
// Access the last paragraph
WParagraph paragraph = document.getLastParagraph();
// Append initial text to the paragraph
paragraph.appendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");
// Add an editable range to the paragraph
EditableRangeStart editableRangeStart = paragraph.appendEditableRangeStart();
// Append editable text
paragraph.appendText("sample databases are based, is a large, multinational manufacturing company.");
// End the editable range
paragraph.appendEditableRangeEnd(editableRangeStart);
// Protect the document and allow only reading
document.protect(ProtectionType.AllowOnlyReading, "password");
// Save the document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}  

### Retrieve the ID of an editable range

You can retrieve the ID of an editable range using the **getId()** method of the **EditableRange** class.

The following code example illustrates how to retrieve the ID of an editable range from a Word document.

{% tabs %}  

{% highlight JAVA %}
// Create a Word document
WordDocument document = new WordDocument();
// Ensure the document has at least one section and one paragraph
document.ensureMinimal();
// Access the last paragraph
WParagraph paragraph = document.getLastParagraph();
// Append text before the editable range
paragraph.appendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");
// Append editable range start
EditableRangeStart editableRangeStart = paragraph.appendEditableRangeStart();
// Append editable content
paragraph.appendText("sample databases are based, is a large, multinational manufacturing company.");
// Append editable range end
paragraph.appendEditableRangeEnd(editableRangeStart);
// Retrieve editable range ID
String editableRangeId = editableRangeStart.getId();
// Protect the document to allow only reading
document.protect(ProtectionType.AllowOnlyReading, "password");
// Save the document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

### Find an editable range

You can find an editable range by its ID in the **EditableRangeCollection** using the **findById()** method.

The following code example illustrates how to find the editable range in a Word document.

{% tabs %}  

{% highlight JAVA %}
// Load an existing Word document
WordDocument document = new WordDocument("Template.docx");
// Get the editable range by ID
EditableRange editableRange = document.getEditableRanges().findById("0");
// Save the document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

### Remove an editable range

You can remove an editable range using the **remove()** method of the **EditableRangeCollection** class.

The following code example demonstrates how to remove an editable range from a Word document.

{% tabs %}  

{% highlight JAVA %}
// Load an existing Word document
WordDocument document = new WordDocument("Template.docx");
// Get the editable range by ID
EditableRange editableRange = document.getEditableRanges().findById("0");
// Remove the editable range
document.getEditableRanges().remove(editableRange);
// Save the Word document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

An editable range at a specific index can also be removed from the **EditableRangeCollection** using the **removeAt()** method.

The following code example demonstrates how to remove an editable range at a particular index from a Word document.

{% tabs %}  

{% highlight JAVA %}
// Load an existing Word document
WordDocument document = new WordDocument("Template.docx");
// Remove the editable range at index 1
document.getEditableRanges().removeAt(1);
// Save the Word document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

### Editing permission

You can restrict editable ranges to specific groups or individuals.

#### Group permission

You can make an editable range accessible to a specific group using the **setEditorGroup()** method of the **EditableRangeStart** class.

The following code example illustrates how to make an editable range available to a group in a Word document.

{% tabs %}  

{% highlight JAVA %}
// Create a new Word document
WordDocument document = new WordDocument();
// Ensure the document has at least one section and one paragraph
document.ensureMinimal();
// Access the last paragraph
WParagraph paragraph = document.getLastParagraph();
// Append text before the editable range
paragraph.appendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");
// Add an editable range start
EditableRangeStart editableRangeStart = paragraph.appendEditableRangeStart();
// Set the editor group to Everyone
editableRangeStart.setEditorGroup(EditorType.Everyone);
// Append text inside the editable range
paragraph.appendText("sample databases are based, is a large, multinational manufacturing company.");
// Add the editable range end
paragraph.appendEditableRangeEnd(editableRangeStart);
// Protect the document with a password, allowing only reading
document.protect(ProtectionType.AllowOnlyReading, "password");
// Save the Word document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();     
{% endhighlight %}

{% endtabs %}

#### Single user permission

Use the **setSingleUser()** method to assign editing permissions to a specific user.

The following code example illustrates how to make an editable range available to a single user in a Word document.

{% tabs %}  

{% highlight JAVA %}
// Create a new Word document
WordDocument document = new WordDocument();
// Ensure the document has at least one section and one paragraph
document.ensureMinimal();
// Access the last paragraph
WParagraph paragraph = document.getLastParagraph();
// Append text before the editable range
paragraph.appendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");
// Add an editable range start
EditableRangeStart editableRangeStart = paragraph.appendEditableRangeStart();
// Set the single user allowed to edit this range
editableRangeStart.setSingleUser("user@domain.com");
// Append text inside the editable range
paragraph.appendText("sample databases are based, is a large, multinational manufacturing company.");
// Add the editable range end
paragraph.appendEditableRangeEnd(editableRangeStart);
// Protect the document with a password, allowing only reading
document.protect(ProtectionType.AllowOnlyReading, "password");
// Save the Word document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

### Add an editable range in a table

Using the **setFirstColumn()** and **setLastColumn()** methods of the **EditableRangeStart** class, you can specify the starting and ending columns of an editable range within a table.

The following code example illustrates how to add an editable range inside a table in a Word document.

{% tabs %}  

{% highlight JAVA %}
// Load an existing Word document
WordDocument document = new WordDocument("Data/Template.docx");
// Access the first table in the first section
WTable table = (WTable) document.getSections().get(0).getTables().get(0);
// Access the paragraph in the 3rd row and 3rd column (index 2,2)
WParagraph paragraph = (WParagraph) table.getRows().get(2).getCells().get(2).getChildEntities().get(0);
// Create a new editable range start for the paragraph
EditableRangeStart editableRangeStart = new EditableRangeStart(document);
// Insert the editable range start at the beginning of the paragraph
paragraph.getChildEntities().insert(0, editableRangeStart);
// Set the editor group to allow everyone to edit
editableRangeStart.setEditorGroup(EditorType.Everyone);
// Apply the editable range to the second column only (index 1)
editableRangeStart.setFirstColumn((short) 0);
editableRangeStart.setLastColumn((short) 1);
// Access another paragraph in the 6th row, 3rd column (index 5,2)
paragraph = (WParagraph) table.getRows().get(5).getCells().get(2).getChildEntities().get(0);
// Append an editable range end to close the region
paragraph.appendEditableRangeEnd();
// Protect the document with a password and allow only reading
document.protect(ProtectionType.AllowOnlyReading, "password");
// Save the Word document
document.save("EditableRange.docx", FormatType.Docx);
// Close the document
document.close();
{% endhighlight %}

{% endtabs %}

N> 1. Editable ranges are supported only in DOCX format.
N> 2. The **SingleUser** and **EditorGroup** properties cannot be set simultaneously for the same editable range. Setting one will clear the other.
