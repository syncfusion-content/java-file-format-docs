---
title: Configure Gradle for Syncfusion Java packages | Syncfusion
description: Learn how to configure Gradle repositories and dependencies to download and use Syncfusion Java packages
platform: java-file-formats
control: general
documentation: UG
---
# Configure Gradle for Syncfusion<sup style="font-size:70%">&reg;</sup> Java packages

You can easily download the Syncfusion<sup style="font-size:70%">&reg;</sup> packages for Java using the [maven repository](https://jars.syncfusion.com/).
 
The following command shows how to mention the repository in Gradle.

{% tabs %}
{% highlight gradle tabtitle="Gradle" %}
repositories {
    maven {
       //Syncfusion® maven repository to download the artifacts.
       url "https://jars.syncfusion.com/repository/maven-public/"
    }
}
{% endhighlight %}
{% endtabs %}

The following command shows how to refer to the Syncfusion<sup style="font-size:70%">&reg;</sup> package in Gradle, which needs to be used in your project as the dependency.

{% tabs %}
{% highlight gradle tabtitle="Gradle" %}
dependencies {
    implementation 'com.syncfusion:syncfusion-docio:18.4.0.30'
}
{% endhighlight %}
{% endtabs %}
