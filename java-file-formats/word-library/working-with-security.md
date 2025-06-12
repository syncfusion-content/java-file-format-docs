---
title: Working with Security | Word library | Syncfusion
description: This section illustrates how to encrypt, decrypt and protect the Word document using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Security

You can encrypt a Word document with password to restrict unauthorized access. You can also control the types of changes you make to this document.

## Encrypting with password

The following code example shows how to encrypt the Word document with password.

{% tabs %}  

{% highlight JAVA %}
//Open an input Word document.
WordDocument document = new WordDocument("Template.docx");
//Encrypt the Word document with a password.
document.encryptDocument("password");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Opening the encrypted Word document

The following code example shows how to open the encrypted Word document. 

{% tabs %}  

{% highlight JAVA %}
//Open an input Word document.
WordDocument document = new WordDocument("Template.docx","password");
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}

## Remove encryption

You can open the encrypted Word document and remove the encryption from the document. The following code example shows how to remove the encryption from encrypted Word document.

{% tabs %}  

{% highlight JAVA %}
//Open an encrypted Word document.
WordDocument document = new WordDocument ("Template.docx", "password");
//Remove encryption in Word document.
document.removeEncryption();
//Save and close the Word document instance.
document.save("Sample.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Protecting Word document from editing

You can restrict a Word document from editing either by providing a password or without password. 

The following are the types of protection:

1. `AllowOnlyComments`: You can add/modify only the comments in the Word document.

2. `AllowOnlyFormFields`: You can modify the form field values in the Word document.

3. `AllowOnlyRevisions`: You can accept or reject the revisions in the Word document.

4. `AllowOnlyReading`: You can only view the content in the Word document.

5. `NoProtection`: You can access/edit the Word document contents as normally.

The following code example shows how to restrict editing to modify only form fields in a Word document.

{% tabs %}  

{% highlight JAVA %}

//Open a Word document.
WordDocument document = new WordDocument("Template.docx");
//Set the protection with password and it allows only to modify the form fields type.
document.protect(ProtectionType.AllowOnlyFormFields, "password"); 
//Save the Word document.
document.save("Protection.docx", FormatType.Docx);
document.close();
{% endhighlight %}

{% endtabs %}  

## Editable ranges

An **editable range** is a portion of a Word document that allows editing even when the document is protected. In the Syncfusion .NET Word library (DocIO), editable ranges are represented using the **EditableRange** class. You can define these ranges programmatically to allow user edits within protected documents.

### Add an editable range

You can add an editable range in the Word document by using **AppendEditableRangeStart** and **AppendEditableRangeEnd** methods of [WParagraph](https://help.syncfusion.com/cr/document-processing/Syncfusion.DocIO.DLS.WParagraph.html) class.

The following code example illustrates how to add an editable range in the Word document.

N> 1. DocIO supports editable ranges in DOCX format documents only.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}  

### Retrieve Id of an editable range

You can retrieve the ID of an editable range using the **Id** property of the **EditableRange** class. 

The following code example illustrates how to retrieve the ID of an editable range from a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

### Find an editable range

You can find an editable range of specific id in the collection of editable ranges through **FindById** method of **EditableRangeCollection** class. 

The following code example illustrates how to find the editable range in a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

### Remove an editable range

You can remove an editable range using the **Remove** method of the **EditableRangeCollection** class.

The following code example demonstrates how to remove an editable range from a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

An editable range at a specific index can also be removed from the **EditableRangeCollection** using the **RemoveAt** method.

The following code example demonstrates how to remove an editable range at particular index from a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

### Editing permission

You can restrict editable ranges to specific groups or individuals.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

#### Group permission

You can make an editable range editable by a group using the **EditorGroup** property of the **EditableRange** class.

The following code example illustrates how to make an editable range available to a group in a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

#### Single user permission

Use the **SingleUser** property of the **EditableRange** class to make an editable range available to a single user for editing.

The following code example illustrates how to make an editable range available to a single user in a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

### Add editable range in a table

Using the **FirstColumn** and **LastColumn** properties of the **EditableRange** class, you can specify the starting and ending columns of an editable range within a table.

The following code example illustrates how to add an editable range inside a table in a Word document.

{% tabs %}  

{% highlight JAVA %}
{% endhighlight %}

{% endtabs %}

N> 1. Editable ranges are supported only in DOCX format.
N> 2. The **SingleUser** and **EditorGroup** properties cannot be set simultaneously for the same editable range. Setting one will clear the other.