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
   <id>Syncfusion-Java</id>
   <name>Syncfusion Java repo</name>
   <url>https://jars.syncfusion.com/repository/maven-public/</url>
</repository>
{% endhighlight %}

{% endtabs %}

The following command shows how to refer to the Syncfusion package, which needs to be used in your project as the dependency.

{% tabs %}  

{% highlight XML %}
<dependency>
   <groupId>com.syncfusion</groupId>
   <artifactId>syncfusion-docio</artifactId>
   <version>18.4.0.30</version>
</dependency>
{% endhighlight %}

{% endtabs %}
