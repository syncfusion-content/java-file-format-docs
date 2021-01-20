---
title: FAQ of Java Word library (DocIO) | Syncfusion
description: This section illustrates about Frequently Asked Questions in Syncfusion Java Word library (Essential DocIO)
platform: java-file-formats
control: Word Library
documentation: UG
---
# Frequently Asked Questions

The frequently asked questions in Syncfusion Java Word library are listed below.

## How to configure Syncfusion Java packages in Gradle?
You can easily download the Syncfusion packages for Java using the [maven repository](https://jars.syncfusion.com/).
 
The following command shows how to mention the repository in Gradle.

<table>
<tr>
<td>
repositories {
maven {
// Syncfusion maven repository to download the artifacts
url "https://jars.syncfusion.com/repository/maven-public/"
}
}
</td>
</tr>
</table>

The following command shows how to refer to the Syncfusion package in Gradle, which needs to be used in your project as the dependency.

<table>
<tr>
<td>
dependencies {
implementation 'com.syncfusion:syncfusion-javahelper:18.4.0.30'
}
</td>
</tr>
</table>

## How to configure Syncfusion Java packages in Apache Maven?

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
<artifactId>syncfusion-javahelper</artifactId>
<version>18.4.0.30</version>
</dependency>
{% endhighlight %}

{% endtabs %}