---
title: Download Syncfusion Java packages from Apache Maven | Syncfusion
description: This section demonstrates how to configure and download required Jars from Apache Maven (Jar configuration)
platform: java-file-formats
control: general
documentation: UG
---
# Configure to download Syncfusion Java packages from Apache Maven

You can easily download the Syncfusion packages for Java using the [maven repository](https://jars.syncfusion.com/).

The following command shows how to mention the repository in Apache Maven.

{% tabs %}  

{% highlight XML %}
<repository>
&nbsp;&nbsp;&nbsp;<id>Syncfusion-Java</id>
&nbsp;&nbsp;&nbsp;<name>Syncfusion Java repo</name>
&nbsp;&nbsp;&nbsp;<url>https://jars.syncfusion.com/repository/maven-public/</url>
</repository>
{% endhighlight %}

{% endtabs %}

The following command shows how to refer to the Syncfusion package, which needs to be used in your project as the dependency.

{% tabs %}  

{% highlight XML %}
<dependency>
&nbsp;&nbsp;&nbsp;<groupId>com.syncfusion</groupId>
&nbsp;&nbsp;&nbsp;<artifactId>syncfusion-docio</artifactId>
&nbsp;&nbsp;&nbsp;<version>18.4.0.30</version>
</dependency>
{% endhighlight %}

{% endtabs %}
