---
layout: post
title: Overview of Syncfusion license and unlock keys - Syncfusion
description: Learn here about the Syncfusion license and unlock keys and difference between license and unlock keys.
platform: java-file-formats
control: Essential Studio
documentation: ug
---


# Syncfusion Licensing Overview

Starting from v19.1.0.x, if you reference Syncfusion Java packages from trial installer or from [maven repository](https://jars.syncfusion.com) you must also include the Java platforms license key in your projects for the corresponding version.

## Difference between unlock key and license key

Please note that this license key is different from the installer unlock key that you might have used in the past and needs to be separately generated from Syncfusion website. Refer [this](https://www.syncfusion.com/kb/8950/difference-between-the-unlock-key-and-licensing-key) KB article to know more about difference between the Syncfusion Unlock Key and the Syncfusion License Key.

Trial message will be displayed as watermark in the generated documents, if Java packages referred from trial installer or from [maven repository](https://jars.syncfusion.com)

**Example**

![IO Licensing Message](licensing-images/io-licensing-message.png)

## Registering Syncfusion license keys in Build server

| Source of Syncfusion assemblies | Details | License Key needs to be registered? | Where to get license key from |
| ------------- | ------------- | ------------- | ------------- |
| **NuGet package** | If the Syncfusion assemblies used in Build Server were from the Syncfusion NuGet packages, then no need to install any Syncfusion installer. We can directly use the required Syncfusion NuGet packages at [nuget.org](http://nuget.org/). <br><br>But, if using NuGet packages from the [nuget.org](https://www.nuget.org/packages?q=syncfusion), then we should register the Syncfusion license key in the application.| Yes | Use any developer license to [generate](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate) keys for Build Environments as well. |
| **Trial installer** | If the Syncfusion assemblies used in Build Server were from Trial Installer, we should register the license key in the application for the corresponding version and platforms, to avoid trial license warning. | Yes | Use any developer trial license to [generate](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate) keys for Build Environments as well. |
| **Licensed installer** |If the Syncfusion assemblies used in Build Server were from Licensed Installer, then there is no need to register the license keys.<r><br>You can [download](https://help.syncfusion.com/java-file-formats/installation/web-installer/how-to-download#download-the-license-version) and [install](https://help.syncfusion.com/java-file-formats/installation/web-installer/how-to-install) the licensed version of our installer. | No | Not applicable |

## See Also

* [How to Generate Syncfusion FileFormats License Key?](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate)
* [How to Register Syncfusion License Key in FileFormats Application?](https://help.syncfusion.com/java-file-formats/licensing/how-to-register-in-an-application)
