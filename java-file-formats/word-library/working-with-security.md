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
