---
title: Loading & Saving document | Syncfusion
description: This section illustrates how to load and save the Word document using Syncfusion Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Loading & saving document

## Opening an existing document

You can open an existing Word document by using either the `open` method or the constructor of `WordDocument` class.

{% tabs %}  

{% highlight JAVA %}
//Open an existing document from the file system using the constructor of WordDocument class.
WordDocument document = new WordDocument(fileName);
{% endhighlight %}

{% endtabs %}

{% tabs %} 
 
{% highlight JAVA %}
//Create an empty Word document instance.
WordDocument document = new WordDocument();
//Load or open an existing word document using the open method of WordDocument class.
document.open(fileName);
{% endhighlight %}

{% endtabs %}  

## Opening an existing document from Stream

You can open an existing document from the stream by using either the overload of `open` methods or the constructor of `WordDocument` class.

{% tabs %}  

{% highlight JAVA %}
//Open an existing document from the stream using the constructor of the WordDocument class.
FileInputStream fileStreamPath = new FileInputStream("Input.docx");
WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic);
{% endhighlight %}

{% endtabs %}  

{% tabs %}   

{% highlight JAVA %}
//Create an empty WordDocument instance.
FileInputStream fileStreamPath = new FileInputStream("Input.docx");
//Loads or opens an existing Word document using the open method of WordDocument class
WordDocument document = new WordDocument();
document.open(fileStreamPath, FormatType.Automatic);
{% endhighlight %}

{% endtabs %}  

## Opening the read only Word document

You can open the ready only documents or read-only streams using the openReadOnly method. If the Word document for reading is opened by any other application such as Microsoft Word, then the same document can be opened using the DocIO in ReadOnly mode. The following code sample explains the same.

{% tabs %}  
  
{% highlight JAVA %}
//Create an empty WordDocument instance.
WordDocument document = new WordDocument();
//Load or open an existing word document using the read only stream.
document.openReadOnly("Template.docx", FormatType.Docx);
{% endhighlight %} 

{% endtabs %}

## Saving a Word document to file system

You can save the created or manipulated Word document to the file system using the `save` method of `WordDocument` class.

{% tabs %}  

{% highlight JAVA %}
//Create an empty WordDocument instance.
WordDocument document = new WordDocument();
//open an existing Word document using the Open method of WordDocument class.
document.open(fileName);
//To-Do some manipulation
//To-Do some manipulation
//Save the document in the file system.
document.save(outputFileName, FormatType.Docx);
{% endhighlight %}

{% endtabs %} 

## Saving a Word document to Stream

You can also save the created or manipulated word document to the stream by using the overloads of `save` methods.

{% tabs %} 

{% highlight JAVA %}
//Creates an empty WordDocument instance
WordDocument document = new WordDocument();
//Opens an existing Word document through Open method of WordDocument class.
document.Open(fileName);
//To-Do some manipulation
//To-Do some manipulation
//Create an instance of the output stream.
ByteArrayOutputStream stream = new ByteArrayOutputStream();
//Save the document to stream.
document.save(stream, FormatType.Docx);
{% endhighlight %}

{% endtabs %}  

## Sending to a client browser

You can save and send the document to a client browser from a web site or web application by invoking the following shown overload of `save` method. This method explicitly makes use of an instance of [HttpResponse](https://msdn.microsoft.com/en-us/library/system.web.httpresponse(v=vs.110).aspx#) as its parameter to stream the document to the client browser. So, this overload is suitable for a web application that references System.Web assembly.

{% tabs %}  

{% highlight JAVA %}
//Create an empty WordDocument instance.
WordDocument document = new WordDocument();
//Open an existing Word document using the Open method of WordDocument class.
document.Open(fileName);
//To-Do some manipulation.
//To-Do some manipulation.
//Create an instance of the output stream.
ByteArrayOutputStream stream = new ByteArrayOutputStream();
//Save the document to the stream.
document.save(outputFileName, FormatType.Docx, Response, HttpContentDisposition.Attachment);
{% endhighlight %}

{% endtabs %}  

## Closing a document

Once the document manipulation and save operation are completed, you should close the instance of `WordDocument`, in order to release all the memory consumed by the DocIO’s DOM. The following code example shows how to close a WordDocument instance.

{% tabs %}  

{% highlight JAVA %}
//Create an empty WordDocument instance.
WordDocument document = new WordDocument();
//Opens an existing Word document using the Open method of WordDocument class
document.Open(fileName);
//To-Do some manipulation.
//To-Do some manipulation.
//Create an instance of the output stream.
ByteArrayOutputStream stream = new ByteArrayOutputStream();
//Save the document to the stream.
document.save(stream, FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}