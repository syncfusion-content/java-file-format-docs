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
&lt;repository&gt;
&lt;id&gt;Syncfusion-Java&lt;/id&gt;
&lt;name&gt;Syncfusion Java repo&lt;/name&gt;
&lt;url&gt;https://jars.syncfusion.com/repository/maven-public/&lt;/url&gt;
&lt;/repository&gt;
{% endhighlight %}

{% endtabs %}

The following command shows how to refer to the Syncfusion package, which needs to be used in your project as the dependency.

{% tabs %}  

{% highlight XML %}
&lt;dependency&gt;
&lt;groupId&gt;com.syncfusion&lt;/groupId&gt;
&lt;artifactId&gt;syncfusion-javahelper&lt;/artifactId&gt;
&lt;version&gt;18.4.0.30&lt;/version&gt;
&lt;/dependency&gt;
{% endhighlight %}

{% endtabs %}