---
title: Working with Form Fields | Word library | Syncfusion
description: This section illustrated how to work with FormFields in Word document using Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Form Fields

You can create template document with form fields such as Text, Checkbox and Drop-Down. You can also open an existing template document and fill the form fields with the specified data. 

The following are the types of form field in the Word document

* Checkbox – represented by an instance of WCheckBox
* Drop-down – represented by an instance of WDropDownFormField
* Text input – represented by an instance of WTextFormField


## Check Box

You can add new Checkbox form field to a Word document by using `appendCheckBox` method of `WParagraph` class.

The following code illustrates how to add new checkbox form field.

{% tabs %}  

{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds new section to the document.
IWSection section = document.addSection();
//Adds new paragraph to the section.
WParagraph paragraph = (WParagraph)section.addParagraph();
paragraph.appendText("Gender\t");
//Appends new Checkbox.
WCheckBox checkbox = paragraph.appendCheckBox();
checkbox.setChecked(false);
//Sets Checkbox size.
checkbox.setCheckBoxSize(10);
checkbox.setCalculateOnExit(true);
//Sets help text.
checkbox.setHelp("Help text");
paragraph.appendText("Male\t");
checkbox = paragraph.appendCheckBox();
checkbox.setChecked(false);
checkbox.setCheckBoxSize(10);
checkbox.setCalculateOnExit(true);
paragraph.appendText("Female");
//Saves the Word document.
document.save("Checkbox.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}  

You can modify the checkbox properties such as checked state, size, help text in a Word document. The following code illustrates how to modify the checkbox form field properties.

{% tabs %} 

{% highlight JAVA %}
//Loads the template document.
WordDocument document = new WordDocument("Checkbox.docx");
//Iterates through paragraph items.
for (Object item_tempObj : document.getLastParagraph().getChildEntities()) 
{
	ParagraphItem item = (ParagraphItem) item_tempObj;
	if (item instanceof WCheckBox) 
	{
		WCheckBox checkbox = (WCheckBox) item;
		//Modifies check box properties.
		if (checkbox.getChecked())
			checkbox.setChecked(false);
		checkbox.setSizeType(CheckBoxSizeType.Exactly);
	}
}
//Saves the Word document.
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}  

## Drop-Down

You can add new Dropdown form field to a Word document by using `appendDropDownFormField` method of `WParagraph` class.

The following code illustrates how to add a new dropdown field.

{% tabs %}  

{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds new section to the document.
IWSection section = document.addSection();
//Adds new paragraph to the section.
WParagraph paragraph = (WParagraph)section.addParagraph();
paragraph.appendText("Educational Qualification\t");
//Appends Dropdown field.
WDropDownFormField dropDownField = paragraph.appendDropDownFormField();
//Adds items to the Dropdown items collection.
dropDownField.getDropDownItems().add("Higher");
dropDownField.getDropDownItems().add("Vocational");
dropDownField.getDropDownItems().add("Universal");
dropDownField.setEnabled(true);
//Sets the item index for default value.
dropDownField.setDropDownSelectedIndex(1);
dropDownField.setCalculateOnExit(true);
//Saves the Word document.
document.save("Dropdown.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}  

You can add or modify list of items of a Dropdown form field in a Word document. The following code illustrates how to modify the dropdown list of a Dropdown form field.

{% tabs %}  

{% highlight JAVA %}
//Loads the template document.
WordDocument document = new WordDocument("Dropdown.docx");
//Iterates through paragraph items.
for (Object item_tempObj : document.getLastParagraph().getChildEntities()) 
{
	ParagraphItem item = (ParagraphItem) item_tempObj;
	if (item instanceof WDropDownFormField) 
	{
		WDropDownFormField dropdown = (WDropDownFormField)item;
		//Modifies the dropdown items.
		dropdown.getDropDownItems().remove(1);
		dropdown.setDropDownSelectedIndex(0);
		dropdown.getCharacterFormat().setFontName("Arial");
	}
}
//Saves the Word document.
document.save("Sample.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}  

## Text Form field

You can add new text form field to a Word document by using `appendTextFormField` method of `WParagraph` class.

The following code illustrates how to add new text form field.

{% tabs %} 

{% highlight JAVA %}
//Creates a new Word document.
WordDocument document = new WordDocument();
//Adds new section to the document.
IWSection section = document.addSection();
//Adds new paragraph to the section.
WParagraph paragraph = (WParagraph)section.addParagraph();
paragraph.appendText("General Information");
section.addParagraph();
paragraph = (WParagraph)section.addParagraph();
IWTextRange text = paragraph.appendText("Name\t");
text.getCharacterFormat().setBold(true);
//Appends Text form field.
WTextFormField textField = paragraph.appendTextFormField(null);
//Sets type of Text form field.
textField.setType(TextFormFieldType.RegularText);
textField.getCharacterFormat().setFontName("Calibri");
textField.setCalculateOnExit(true);
section.addParagraph();
paragraph = (WParagraph)section.addParagraph();
text = paragraph.appendText("Date of Birth\t");
text.getCharacterFormat().setBold(true);
//Appends Text form field.
textField = paragraph.appendTextFormField("Date field", DateTimeSupport.toString(LocalDateTime.now(), "MM/DD/YY"));
textField.setStringFormat("MM/DD/YY");
textField.setType(TextFormFieldType.DateText);
textField.setCalculateOnExit(true);
//Saves the Word document.
document.save("TextForm.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}  

You can add or modify text form field properties such as default text, type in a Word document. The following code illustrates how to modify the text form field

{% tabs %} 

{% highlight JAVA %}
//Loads the template document. 
WordDocument document = new WordDocument("TextForm.docx");
//Iterates through section.
for (Object section_tempObj : document.getSections()) 
{
	WSection section = (WSection) section_tempObj;
	//Iterates through section child elements.
	for (Object textBody_tempObj : section.getChildEntities()) 
	{
		WTextBody textBody = (WTextBody) textBody_tempObj;
		//Iterates through form fields.
		for (Object formField_tempObj : textBody.getFormFields())
		{
			WFormField formField = (WFormField) formField_tempObj;
			switch (formField.getFormFieldType().toString()) 
			{
				case "TextInput":
					WTextFormField textField = (WTextFormField) formField;
					if (textField.getType().getEnumValue() == TextFormFieldType.DateText.getEnumValue()) 
					{
						//Modifies the text form field.
						textField.setType(TextFormFieldType.RegularText);
						textField.setStringFormat("");
						textField.setDefaultText("Default text");
						textField.setText("Default text");
						textField.setCalculateOnExit(false);
					}
					break;
			}
		}
	}
}
//Saves the Word document.
document.save("Sample.docx", FormatType.Docx);
//Closes the document.
document.close();
{% endhighlight %}

{% endtabs %}