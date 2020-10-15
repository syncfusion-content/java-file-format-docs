---
title: Working with Mail merge | Syncfusion
description: This section illustrates about Mail merge Word document to create reports (letters, envelopes, labels, invoice, payroll) without MS Word or Office interop.
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Mail merge

Mail merge is a process of merging data from data source to a Word template document. The `WMergeField` class provides support to bind template document and data source. The `WMergeField` instance is replaced with the actual data retrieved from data source for the given merge field name in a template document.

The following data sources are supported by Essential DocIO for performing Mail merge:

* String Arrays
* Java objects

## Mail merge process

The mail merge process involves three documents:

1. **Template Word document**: This document contains the static or templated text and graphics along with the merge fields (that are placeholders) for replacing dynamic data.

2. **Data source**: This represents file containing data to replace the merge fields in template Word document.

3. **Final merged document**: This resultant document is a combination of the template Word document and the data from data source.

T> 1. You can use conditional fields ([IF](https://support.office.com/en-us/article/Field-codes-IF-field-9f79e82f-e53b-4ff5-9d2c-ae3b22b7eb5e), [Formula](https://support.office.com/en-us/article/field-codes-formula-field-32d5c9de-3516-4ec3-80ed-d1fc2b5bc21d?ui=en-US&rs=en-US&ad=US)) combined with merge fields, when you require intelligent decisions in addition to simple mail merge (replace merge fields with result text). To use conditional fields, execute mail merge and then update fields in the Word document using `updateDocumentFields` API.
T> 2. You can replace the fields ([IF](https://support.office.com/en-us/article/Field-codes-IF-field-9f79e82f-e53b-4ff5-9d2c-ae3b22b7eb5e), [Formula](https://support.office.com/en-us/article/field-codes-formula-field-32d5c9de-3516-4ec3-80ed-d1fc2b5bc21d?ui=en-US&rs=en-US&ad=US)) combined with merge fields, with its most recent result and **generates the plain Word document** by unlinking the fields. Refer to this [link](https://help.syncfusion.com/java-file-formats/word-library/working-with-fields#unlink-fields) for more information. 

### Create Word document template

You can create a template document with merge fields by using any Word editor application, like Microsoft Word. By using Word editor application, you can take the advantage of the visual interface to design unique layout, formatting, and more for your Word document template interactively. 

The following screenshot shows how to insert a merge field in the Word document by **using the Microsoft Word.**

![Word template document](MailMerge_images/MailMerge_template.png)

You need to add a prefix (“Image:”) to the merge field name for merging an image in the place of a merge field.

**For example:** The merge field name should be like “Image:Photo” (<<Image:MergeFieldName>>)

You can **create Word document template programmatically** by adding merge fields to the Word document using Essential DocIO.

The following code example shows how to create a merge field in the Word document.

{% tabs %}  

{% highlight JAVA %}
//Creates an instance of a WordDocument.
WordDocument document = new WordDocument();
//Adds a section and a paragraph in the document.
document.ensureMinimal();
//Appends merge field to the last paragraph.
document.getLastParagraph().appendField("FullName", FieldType.FieldMergeField);
//Saves the Word document. 
document.save("Template.docx", FormatType.Docx);
//Closes the Word document.
document.close();
{% endhighlight %}
{% endtabs %}

### Execute mail merge

The following code example shows how to perform mail merge in above Word document template using string arrays as data source.

{% tabs %}  
{% highlight JAVA %}
//Opens the template document.
FileInputStream fileStreamPath = new FileInputStream("Template.docx");
WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);
String[] fieldNames = new String[] { "FullName" };
String[] fieldValues = new String[] { "Nancy Davolio" };
//Performs the mail merge.
document.getMailMerge().execute(fieldNames, fieldValues);
//Saves the Word document.
document.save("Output.docx", FormatType.Docx);
//Closes the Word document.
document.close();
{% endhighlight %}
{% endtabs %}

By executing the previous code example, it generates the resultant Word document as follows.

![Mail merge Word document](MailMerge_images/MailMerge_output.png)

## Simple Mail merge

The `MailMerge` class provides various overloads for the `Execute` method to perform Mail merge from various data sources. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/simple-mail-merge). 

## Performing Mail merge for a group

You can perform Mail merge and append multiple records from data source within a specified region to a template document. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-for-group).

## Performing Nested Mail merge for group

You can perform nested Mail merge with relational or hierarchical data source and independent data tables in a template document. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-for-nested-groups).

## Performing Mail merge with business objects

You can perform Mail merge with business objects in a template document. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-for-group#mail-merge-with-Java-objects).

## Performing Nested Mail merge with relational data objects

Essential DocIO supports performing nested Mail merge with implicit relational data objects without any explicit relational commands by using the `ExecuteNestedGroup` overload method. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-for-nested-groups#mail-merge-with-implicit-relational-data).

## Event support for mail merge

The `MailMerge` class provides event support to customize the document contents and merging image data during the Mail merge process. The following events are supported by Essential DocIO in Mail merge process:

* `MergeField`: Occurs when a **Mail merge field** except image Mail merge field is encountered.

* `MergeImageField`: Occurs when an **image Mail merge field** is encountered.

* `BeforeClearGroupField`: Occurs when an **unmerged group field** is encountered.

### MergeField event

You can customize the merging text during Mail merge process by using the `MergeField` event. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-events#mergefield-event).

### MergeImageField event

You can customize the merging image during Mail merge process by using the `MergeImageField` event. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-events#mergeimagefield-event).

### BeforeClearGroupField event

You can get the unmerged groups during Mail merge process by using the `BeforeClearGroupField` event. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-events#beforecleargroupfield-event).

## Mail merge options

The `MailMerge` class allows you to customize the Mail merge process with the following options:

### Field mapping

You can automatically map the merge field names with data source column names during Mail merge process. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-options#field-mapping).

### Retrieving the merge field names

You can retrieve the merge field names and also merge field group names in the Word document. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-options#retrieve-the-merge-field-names).

### Removing empty paragraphs

You can remove the empty paragraphs when the paragraph has a merge field item without any data during Mail merge process. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-options#remove-empty-paragraphs).

### Removing empty merge fields

You can remove or keep the unmerged merge fields in the output document based on the `ClearFields` property on each mail merge execution. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-options#remove-empty-merge-fields).

### Restart numbering in lists

You can restart the list numbering in a Word document during Mail merge. For further information, click [here](https://help.syncfusion.com/java-file-formats/word-library/mail-merge/mail-merge-options#restart-numbering-in-lists).
