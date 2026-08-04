---
title: Download Syncfusion Java packages from Gradle | Syncfusion
description: This section demonstrates how to configure and download required Jars from Gradle (Jar configuration)
platform: java-file-formats
control: general
documentation: UG
---
# Configure to download Java packages from Gradle

You can easily download the Syncfusion<sup style="font-size:70%">&reg;</sup> packages for Java using the [maven repository](https://jars.syncfusion.com/).
 
The following snippet shows how to add the repository in the `build.gradle` file of your Gradle project.

{% tabs %}
{% highlight groovy tabtitle="Gradle" %}
repositories {
    maven {
       //Syncfusion® maven repository to download the artifacts.
       url "https://jars.syncfusion.com/repository/maven-public/"
    }
}
{% endhighlight %}
{% endtabs %}

The following snippet shows how to add the Syncfusion<sup style="font-size:70%">&reg;</sup> package in the `build.gradle` file, which needs to be used in your project as the dependency.

{% tabs %}
{% highlight groovy tabtitle="Gradle" %}
dependencies {
    implementation 'com.syncfusion:syncfusion-docio:18.4.0.30'
}
{% endhighlight %}
{% endtabs %}

N> The version `18.4.0.30` shown above is for illustration only. Replace it with the [latest Syncfusion<sup style="font-size:70%">&reg;</sup> Java package version](https://jars.syncfusion.com/) available for your license.
