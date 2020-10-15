---
title: Mail merge events | Syncfusion
description: This section illustrates how to format or customize the merged text and image, clear or retain unmerged fields during mail merge using events.
platform: java-file-formats
control: Word Library
documentation: UG
---

# Event support for Mail merge

The `MailMerge` class provides event support to customize the document contents and merging image data during the Mail merge process. The following events are supported by Essential DocIO during Mail merge process:

* `MergeField`- occurs when a **Mail merge field** except image Mail merge field is encountered.

* `MergeImageField`- occurs when an **image Mail merge field** is encountered.

* `BeforeClearGroupField`- occurs when an **unmerged group field** is encountered.

## MergeField Event

You can apply formatting to the merged text or customize the merged text during mail merge process using the `MergeField` Event.

The following code example shows how to use the `MergeField` event during Mail merge process.

{% tabs %}  

{% highlight JAVA %}
//Opens the template document.
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Uses the mail merge events to perform the conditional formatting during runtime.
document.getMailMerge().MergeField.add("applyAlternateRecordsTextColor", new MergeFieldEventHandler() {
ListSupport<MergeFieldEventHandler> delegateList = new ListSupport<MergeFieldEventHandler>(
MergeFieldEventHandler.class);
// Represents event handling for MergeFieldEventHandlerCollection.
public void invoke(Object sender, MergeFieldEventArgs args) throws Exception 
{
	applyAlternateRecordsTextColor(sender, args);
}
// Represents the method that handles MergeField event.
public void dynamicInvoke(Object... args) throws Exception 
{
	applyAlternateRecordsTextColor((Object) args[0], (MergeFieldEventArgs) args[1]);
}
// Represents the method that handles MergeField event to add collection item.
public void add(MergeFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.add(delegate);
}
// Represents the method that handles MergeField event to remove collection item.
public void remove(MergeFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.remove(delegate);
}
});
//Executes Mail Merge with groups.
document.getMailMerge().executeGroup(getDataTable());
//Saves the Word document.
document.save("Sample.docx", FormatType.Docx);
//Closes the Word document.
document.close();
{% endhighlight %} 

{% endtabs %}  

The following code example shows how to set text color to the alternate Mail merge record by using MergeFieldEventHandler.

{% tabs %} 

{% highlight JAVA %}
private void applyAlternateRecordsTextColor (Object sender, MergeFieldEventArgs args) throws Exception
{
    //Sets text color to the alternate mail merge record.
	if (Integer.compare((args.getRowIndex() % 2),0)==0)
	{
		args.getTextRange().getCharacterFormat().setTextColor(ColorSupport.fromArgb(255, 102, 0));
	}
}
{% endhighlight %}

{% endtabs %}  

The following code example shows getDataTable method which are is to get data for mail merge.

{% tabs %} 

{% highlight JAVA %}
private static DataTableSupport getDataTable() throws Exception
{
	DataTableSupport dataTable = new DataTableSupport("Employee");
	dataTable.getColumns().add("EmployeeName");
	dataTable.getColumns().add("EmployeeNumber");
	for (int i = 0; i < 20; i++)
	{
		DataRowSupport datarow = dataTable.newRow();
		dataTable.getRows().add(datarow);
		datarow.set(0 , "Employee" + Integer.toString(i));
		datarow.set(1 , "EMP" + Integer.toString(i));
	}
	return dataTable;
}
{% endhighlight %}

{% endtabs %} 

## MergeImageField Event

You can format the merged image like resizing the image and more during mail merge process using the `MergeImageField` Event. 

The following code example shows how to use the `MergeImageField` event during Mail merge process.

{% tabs %}  

{% highlight JAVA %}
WordDocument document = new WordDocument("Template.docx", FormatType.Docx);
//Uses the mail merge events handler for image fields.
document.getMailMerge().MergeImageField.add("mergeField_ProductImage", new MergeImageFieldEventHandler() {
ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
MergeImageFieldEventHandler.class);
//Represents event handling for MergeImageFieldEventHandlerCollection.
public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception
{
	mergeField_ProductImage(sender, args);
}
//Represents the method that handles MergeImageField event.
public void dynamicInvoke(Object... args) throws Exception 
{
	mergeField_ProductImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
}
//Represents the method that handles MergeImageField event to add collection item.
public void add(MergeImageFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.add(delegate);
}
//Represents the method that handles MergeImageField event to remove collection item.
public void remove(MergeImageFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		elegateList.remove(delegate);
}
});
//Specifies the field names and field values.
String[] fieldNames = new String[] { "Logo"};
String[] fieldValues = new String[] { "Logo.png"};
//Executes the mail merge with groups.
document.getMailMerge().execute(fieldNames, fieldValues);
//Saves the Word document.
document.save("Samples.docx", FormatType.Docx);
//Closes the Word document.
document.close();
{% endhighlight %}

{% endtabs %}  
  
The following code example shows how to bind the image from file system during Mail merge process by using MergeImageFieldEventHandler.

{% tabs %}  

{% highlight JAVA %}
private void mergeField_ProductImage(Object sender, MergeImageFieldEventArgs args) throws Exception
{ 
	//Binds image from file system during mail merge.
	if ((args.getFieldName()).equals("Logo"))
	{
		String ProductFileName = args.getFieldValue().toString();
		//Gets the image from file system.
		FileStreamSupport imageStream = new FileStreamSupport(ProductFileName, FileMode.Open, FileAccess.Read);
		ByteArrayInputStream stream = new ByteArrayInputStream(imageStream.toArray());
		args.setImageStream(stream);
		//Gets the picture, to be merged for image merge field.
		WPicture picture = args.getPicture();
		//Resizes the picture.
		picture.setHeight(50);
		picture.setWidth(150);
	}
}
{% endhighlight %}

{% endtabs %} 


## BeforeClearGroupField Event

You can get the unmerged group fields in a Word document during mail merge process using the `BeforeClearGroupField` event.

The following code example shows how to use the `BeforeClearGroupField` event during Mail merge process.

{% tabs %}  

{% highlight JAVA %}
// Opens the template document.
WordDocument document = new WordDocument("Template.docx",FormatType.Docx);
// Sets “ClearFields” to true to remove empty mail merge fields from document.
document.getMailMerge().setClearFields(false);
// Uses the mail merge event to clear the unmerged group field while perform mail merge execution.
document.getMailMerge().BeforeClearGroupField.add("beforeClearFields", new BeforeClearGroupFieldEventHandler() {
ListSupport<BeforeClearGroupFieldEventHandler> delegateList = new ListSupport<BeforeClearGroupFieldEventHandler>(
BeforeClearGroupFieldEventHandler.class);
// Represents event handling for BeforeClearGroupFieldEvent.
public void invoke(Object sender, BeforeClearGroupFieldEventArgs args) throws Exception 
{
	beforeClearFields(sender, args);
}
// Represents the method that handles BeforeClearGroupField event.
public void dynamicInvoke(Object... args) throws Exception 
{
	beforeClearFields((Object) args[0], (BeforeClearGroupFieldEventArgs) args[1]);
}
// Represents the method that handles BeforeClearGroupField event to add collection item.
public void add(BeforeClearGroupFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.add(delegate);
}
// Represents the method that handles BeforeClearGroupField event to remove collection item.
public void remove(BeforeClearGroupFieldEventHandler delegate) throws Exception 
{
	if (delegate != null)
		delegateList.remove(delegate);
}
});
// Gets the employee details as “IEnumerable” collection.
ListSupport<Employees> employeeList = getEmployees();
// Creates an instance of “MailMergeDataTable” by specifying mail merge group
// name and “IEnumerable” collection.
MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
// Performs Mail merge.
document.getMailMerge().executeNestedGroup(dataTable);
// Saves the Word document.
document.save("Sample.docx", FormatType.Docx);
// Closes the Word document.
document.close();
{% endhighlight %}

{% endtabs %}

The following code example shows how to bind the data to unmerged group fields during Mail merge process by using BeforeClearGroupFieldEventHandler.

{% tabs %}  

{% highlight JAVA %}
private static void beforeClearFields(Object sender, BeforeClearGroupFieldEventArgs args) throws Exception 
{
	if (!args.getHasMappedGroupInDataSource()) 
	{
		// Gets the Current unmerged group name from the event argument.
		String[] groupName = args.getGroupName().split(":");
		if ((groupName[groupName.length - 1]).equals("Orders")) 
		{
			String[] fields = args.getFieldNames();
			ListSupport<OrderDetails> orderList = getOrders();
			// Binds the data to the unmerged fields in group as alternative values.
			args.setAlternateValues(orderList);
		} 
		else
			// If group value is empty, you can set whether the unmerged merge group field can be clear or not.
			args.setClearGroup(true);
	}
}
{% endhighlight %}

{% endtabs %} 

The following code example shows getOrders and getEmployees methods which are used to get data for mail merge.

{% tabs %} 

{% highlight JAVA %}
//Gets order list.
private static ListSupport<OrderDetails> getOrders() throws Exception 
{
	ListSupport<OrderDetails> orders = new ListSupport<OrderDetails>();
	orders.add(new OrderDetails("10952", LocalDateTime.of(2015, 2, 5, 0, 0, 0),
	LocalDateTime.of(2015, 2, 12, 0, 0, 0), LocalDateTime.of(2015, 2, 21, 0, 0, 0)));
	return orders;
}
//Gets employee list.
public static ListSupport<Employees> getEmployees() throws Exception 
{
	// Gets the OrderDetails as “IEnumerable” collection.
	ListSupport<OrderDetails> orders = new ListSupport<OrderDetails>();
	orders.add(new OrderDetails("10835", LocalDateTime.of(2015, 1, 5, 0, 0, 0),
	LocalDateTime.of(2015, 1, 12, 0, 0, 0), LocalDateTime.of(2015, 1, 21, 0, 0, 0)));
	// Gets the CustomerDetails as “IEnumerable” collection.
	ListSupport<CustomerDetails> customerDetails = new ListSupport<CustomerDetails>();
	customerDetails.add(new CustomerDetails("Maria Anders", "Maria Anders", "Berlin", "Germany", orders));
	customerDetails.add(new CustomerDetails("Andy", "Bernard", "Berlin", "Germany", null));
	// Gets the Employees details as “IEnumerable” collection.
	ListSupport<Employees> employees = new ListSupport<Employees>();
	employees.add(new Employees("Nancy", "Smith", "1", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "USA", customerDetails));
	return employees;
}
{% endhighlight %}

{% endtabs %} 

The following code example shows Employees, CustomerDetails, and OrderDetails classes.

{% tabs %}  
{% highlight JAVA %}
public class Employees 
{
	private String _firstName;
	private String _lastName;
	private String _employeeID;
	private String _address;
	private String _city;
	private String _country;
	private ListSupport<CustomerDetails> _customers;
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
	public String getEmployeeID() throws Exception 
	{
		return _employeeID;
	}
	public String setEmployeeID(String value) throws Exception 
	{
		_employeeID = value;
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
	public String getCountry() throws Exception 
	{
		return _country;
	}
	public String setCountry(String value) throws Exception 
	{
		_country = value;
		return value;
	}
	public ListSupport<CustomerDetails> getCustomers() throws Exception 
	{
		return _customers;
	}
	public ListSupport<CustomerDetails> setCustomers(ListSupport<CustomerDetails> value) throws Exception 
	{
		_customers = value;
		return value;
	}

	public Employees(String firstName, String lastName, String employeeId, String address, String city, String country,ListSupport<CustomerDetails> customers) throws Exception 
	{
		setFirstName(firstName);
		setLastName(lastName);
		setAddress(address);
		setEmployeeID(employeeId);
		setCity(city);
		setCountry(country);
		setCustomers(customers);
	}
}

public class CustomerDetails 
{
	private String _contactName;
	private String _companyName;
	private String _city;
	private String _country;
	private ListSupport<OrderDetails> _orders;
	public String getContactName() throws Exception 
	{
		return _contactName;
	}
	public String setContactName(String value) throws Exception 
	{
		_contactName = value;
		return value;
	}
	public String getCompanyName() throws Exception
	{
		return _companyName;
	}
	public String setCompanyName(String value) throws Exception 
	{
		_companyName = value;
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
	public String getCountry() throws Exception 
	{
		return _country;
	}
	public String setCountry(String value) throws Exception 
	{
		_country = value;
		return value;
	}
	public ListSupport<OrderDetails> getOrders() throws Exception 
	{
		return _orders;
	}
	public ListSupport<OrderDetails> setOrders(ListSupport<OrderDetails> value) throws Exception 
	{
		_orders = value;
		return value;
	}
	public CustomerDetails(String contactName, String companyName, String city, String country,ListSupport<OrderDetails> orders) throws Exception 
	{
		setContactName(contactName);
		setCompanyName(companyName);
		setCity(city);
		setCountry(country);
		setOrders(orders);
	}
}

public class OrderDetails 
{
	private String _orderID;
	private LocalDateTime _orderDate;
	private LocalDateTime _shippedDate;
	private LocalDateTime _requiredDate;
	public String getOrderID() throws Exception 
	{
		return _orderID;
	}
	public String setOrderID(String value) throws Exception 
	{
		_orderID = value;
		return value;
	}
	public LocalDateTime getOrderDate() throws Exception 
	{
		return _orderDate;
	}
	public LocalDateTime setOrderDate(LocalDateTime value) throws Exception 
	{
		_orderDate = value;
		return value;
	}
	public LocalDateTime getShippedDate() throws Exception 
	{
		return _shippedDate;
	}
	public LocalDateTime setShippedDate(LocalDateTime value) throws Exception 
	{
		_shippedDate = value;
		return value;
	}
	public LocalDateTime getRequiredDate() throws Exception 
	{
		return _requiredDate;
	}
	public LocalDateTime setRequiredDate(LocalDateTime value) throws Exception 
	{
		_requiredDate = value;
		return value;
	}
	public OrderDetails(String orderId, LocalDateTime orderDate, LocalDateTime shippedDate, LocalDateTime requiredDate) throws Exception 
	{
		setOrderID(orderId);
		setOrderDate(orderDate);
		setShippedDate(shippedDate);
		setRequiredDate(requiredDate);
	}
}
{% endhighlight %}
{% endtabs %}
