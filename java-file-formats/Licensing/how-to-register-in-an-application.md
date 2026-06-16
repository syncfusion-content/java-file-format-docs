---
layout: post
title: Overview of Syncfusion license registration - Syncfusion
description: Learn here about how to register Syncfusion FileFormat license key for FileFormat application for license validation.
platform: java-file-formats
control: Essential Studio
documentation: ug
---

# Register Syncfusion<sup style="font-size:70%">&reg;</sup> License Key in FileFormat Application

The generated license key is just a string that needs to be registered before any Syncfusion<sup style="font-size:70%">&reg;</sup> control is initiated. The following code is used to register the license.

{% tabs %}
{% highlight c# %}
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR LICENSE KEY");
{% endhighlight %}
{% endtabs %}

N> * Place the license key between double quotes. Also, ensure that Syncfusion.Licensing.dll is referenced in your project where the license key is being registered.
* Syncfusion<sup style="font-size:70%">&reg;</sup> license validation is done offline during application execution and does not require internet access. Apps registered with a Syncfusion<sup style="font-size:70%">&reg;</sup> license key can be deployed on any system that does not have an internet connection.

### Java

The recommended place to register the license for the Java platform is given below.

Import the ‘syncfusion.licensing’ package and register the license key in the **main method** of your console application.

{% tabs %}
{% highlight java %}
// Refer to the licensing package
import com.syncfusion.licensing.*;

static void main() { 
    // Register Syncfusion license 
    SyncfusionLicenseProvider.registerLicense("YOUR LICENSE KEY"); 
}
{% endhighlight %}
{% endtabs %}

N> License key registration is not required for Java versions prior to v19.1.