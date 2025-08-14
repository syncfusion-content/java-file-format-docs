---
title: Mail Merge for Group | Syncfusion
description: This section illustrates how to perform a Mail Merge for a group - replacing merge fields in a region of a document with data by repeating the region for each record.
platform: java-file-formats
control: Word Library
documentation: UG
---

# Mail Merge for a Group

You can perform a Mail Merge and append multiple records from a data source within a specified region to a template document. The region between the start and end group merge fields gets repeated for every record from the data source.

## Create Template for Group Mail Merge

The region where the Mail Merge operations are to be performed must be marked by two merge fields with the following names:

  * «TableStart:TableName» and «BeginGroup:GroupName» - For the entry point of the region.  
  
  * «TableEnd:TableName» and «EndGroup:GroupName» - For the endpoint of the region.  
  
																									  
																												
  
																 

  1. *TableStart* and *TableEnd* regions are preferred for performing Mail Merge inside the table cell.  
  2. *BeginGroup* and *EndGroup* regions are preferred for performing Mail Merge inside the document body contents.  

For example, consider that you have a template document as shown below:

![Mail Merge for a Group](../MailMerge_images/Group_mail_merge_template.png)

In this template, "Employees" is the group name, and the same name should be used while performing Mail Merge through code. There are two special merge fields, "TableStart:Employees" and "TableEnd:Employees," to denote the start and end of the Mail Merge group.

## Mail Merge with Java Objects

You can perform a Mail Merge with Java objects in a template document. The following code snippet shows how to perform a Mail Merge with business objects.

{% tabs %} 

{% highlight JAVA %}
// Loads an existing Word document into the DocIO instance.
WordDocument document = new WordDocument("EmployeesReportDemo.docx");
// Gets the employee details as an IEnumerable collection.
ListSupport<Employee> employeeList = getEmployees();
// Uses the Mail Merge events handler for image fields.
document.getMailMerge().MergeImageField.add("mergeField_EmployeeImage", new MergeImageFieldEventHandler() {
ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
MergeImageFieldEventHandler.class);
// Represents event handling for MergeFieldEventHandlerCollection.
public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception 
{
	mergeField_EmployeeImage(sender, args);
}
// Represents the method that handles the MergeField event.
public void dynamicInvoke(Object... args) throws Exception 
{
	mergeField_EmployeeImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
}
// Represents the method that handles the MergeField event to add a collection item.
public void add(MergeImageFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.add(delegate);
}
// Represents the method that handles the MergeField event to remove a collection item.
public void remove(MergeImageFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.remove(delegate);
}
});
// Creates an instance of MailMergeDataTable by specifying the Mail Merge group name and IEnumerable collection.
MailMergeDataTable dataSource = new MailMergeDataTable("Employees", employeeList);
// Executes the Mail Merge for the group.
document.getMailMerge().executeGroup(dataSource);
// Saves and closes the WordDocument instance.
document.save("Sample.docx");
document.close();
{% endhighlight %}

{% endtabs %}  

The following code example shows the `getEmployees` method, which is used to get data for Mail Merge.

{% tabs %}  

{% highlight JAVA %}
public ListSupport<Employee> getEmployees() throws Exception
{
	ListSupport<Employee> employees = new ListSupport<Employee>(Employee.class);
	employees.add(new Employee("Nancy", "Smith", "Sales Representative", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "WA", "USA", "Nancy.png"));
	employees.add(new Employee("Andrew", "Fuller", "Vice President, Sales", "908 W. Capital Way", "Tacoma", "WA", "USA", "Andrew.png"));
	return employees;
}

{% endhighlight %}

{% endtabs %}  

The following code example shows how to bind the image from the file system during the Mail Merge process by using `MergeImageFieldEventHandler`.

{% tabs %}  

{% highlight JAVA %}
private void mergeField_EmployeeImage(Object sender, MergeImageFieldEventArgs args) throws Exception 
{
	// Binds the image from the file system during Mail Merge.
	if ((args.getFieldName()).equals("Photo")) 
	{
		String ProductFileName = args.getFieldValue().toString();
		// Gets the image from the file system.
		FileStreamSupport imageStream = new FileStreamSupport(ProductFileName, FileMode.Open, FileAccess.Read);
		ByteArrayInputStream stream = new ByteArrayInputStream(imageStream.toArray());
		args.setImageStream(stream);
	}
}

{% endhighlight %}

{% endtabs %}

The following code example shows the `Employee` class.

{% tabs %}  
{% highlight JAVA %}
public class Employee 
{
	private String _firstName;
	private String _lastName;
	private String _address;
	private String _city;
	private String _region;
	private String _country;
	private String _title;
	private String _photo;

	public String getFirstName() throws Exception
	{
		return _firstName;
	}
	public String setFirstName(String value) throws Exception
	{
		_firstName = value;
		return value;
	}
	public String getLastName() throws Exception
	{
		return _lastName;
	}
	public String setLastName(String value) throws Exception
	{
		_lastName = value;
		return value;
	}
	public String getAddress() throws Exception
	{
		return _address;
	}
	public String setAddress(String value) throws Exception
	{
		_address = value;
		return value;
	}
	public String getCity() throws Exception
	{
		return _city;
	}
	public String setCity(String value) throws Exception
	{
		_city = value;
		return value;
	}
	public String getRegion() throws Exception
	{
		return _region;
	}
	public String setRegion(String value) throws Exception
	{
		_region = value;
		return value;
	}
	public String getCountry() throws Exception
	{
		return _country;
	}
	public String setCountry(String value) throws Exception
	{
		_country = value;
		return value;
	}
	public String getTitle() throws Exception
	{
		return _title;
	}
	public String setTitle(String value) throws Exception
	{
		_title = value;
		return value;
	}
	public String getPhoto() throws Exception
	{
		return _photo;
	}
	public String setPhoto(String image) throws Exception
	{
		_photo = image;
		return image;
	}
	public Employee(String firstName, String lastName, String title, String address, String city, String region, String country, String photoFilePath) throws Exception
	{
		setFirstName(firstName);
		setLastName(lastName);
		setTitle(title);
		setAddress(address);
		setCity(city);
		setRegion(region);
		setCountry(country);
		setPhoto(photoFilePath);
	}
}
{% endhighlight %}
			 

{% endtabs %}

The resultant document looks as follows:

![Group Resultant Document](../MailMerge_images/Group_mail_merge_output.png)