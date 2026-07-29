---
title: Download Syncfusion Java packages from Apache Maven | Syncfusion
description: This section demonstrates how to configure and download required Jars from Apache Maven (Jar configuration)
platform: java-file-formats
control: general
documentation: UG
---
# Configure to download Java packages from Apache Maven

You can easily download the Syncfusion<sup style="font-size:70%">&reg;</sup> packages for Java using the [maven repository](https://jars.syncfusion.com/).

The following snippet shows how to add the repository in your Apache Maven project.

{% tabs %}  

{% highlight XML %}
<repository>
   <id>Syncfusion-Java</id>
   <name>Syncfusion<sup style="font-size:70%">&reg;</sup> Java repo</name>
   <url>https://jars.syncfusion.com/repository/maven-public/</url>
</repository>
{% endhighlight %}

{% endtabs %}

The following snippet shows how to add the Syncfusion<sup style="font-size:70%">&reg;</sup> package, which needs to be used in your project as the dependency.

{% tabs %}  

{% highlight XML %}
<dependency>
   <groupId>com.syncfusion</groupId>
   <artifactId>syncfusion-docio</artifactId>
   <version>18.4.0.30</version>
</dependency>
{% endhighlight %}

{% endtabs %}

N> The version `18.4.0.30` shown above is for illustration only. Replace it with the [latest Syncfusion<sup style="font-size:70%">&reg;</sup> Java package version](https://jars.syncfusion.com/) available for your license.
