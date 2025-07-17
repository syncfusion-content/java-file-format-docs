---
layout: post
title: Overview of Syncfusion License and Unlock Keys - Syncfusion
description: Learn about the Syncfusion license and unlock keys, along with the difference between the license and unlock keys.
platform: java-file-formats
control: Essential Studio
documentation: ug
---


# Syncfusion<sup style="font-size:70%">&reg;</sup> Licensing Overview

Starting from v19.1.0.x, if you reference Syncfusion<sup style="font-size:70%">&reg;</sup> Java packages from the trial installer or the [Maven repository](https://jars.syncfusion.com), you must also include the Java platform's license key in your projects for the corresponding version.

## Difference Between Unlock Key and License Key

Please note that this license key is different from the installer unlock key that you might have used in the past and needs to be separately generated from the Syncfusion website. Refer to [this](https://www.syncfusion.com/kb/8950/difference-between-the-unlock-key-and-licensing-key) KB article to learn more about the difference between the Syncfusion<sup style="font-size:70%">&reg;</sup> Unlock Key and the Syncfusion<sup style="font-size:70%">&reg;</sup> License Key.

A trial message will be displayed as a watermark in the generated documents if Java packages are referenced from the trial installer or the [Maven repository](https://jars.syncfusion.com).

**Example**

![IO Licensing Message](licensing-images/io-licensing-message.png)

## Registering Syncfusion<sup style="font-size:70%">&reg;</sup> License Keys in Build Server

| Source of Syncfusion<sup style="font-size:70%">&reg;</sup> Assemblies | Details | License Key Needs to Be Registered? | Where to Get License Key From |
| ------------- | ------------- | ------------- | ------------- |
| **NuGet Package** | If the Syncfusion<sup style="font-size:70%">&reg;</sup> assemblies used in the Build Server were from the Syncfusion<sup style="font-size:70%">&reg;</sup> NuGet packages, there is no need to install any Syncfusion<sup style="font-size:70%">&reg;</sup> installer. We can directly use the required Syncfusion<sup style="font-size:70%">&reg;</sup> NuGet packages from [nuget.org](http://nuget.org). <br><br>However, if using NuGet packages from [nuget.org](https://www.nuget.org/packages?q=syncfusion), then we should register the Syncfusion<sup style="font-size:70%">&reg;</sup> license key in the application. | Yes | Use any developer license to [generate](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate) keys for Build Environments as well. |
| **Trial Installer** | If the Syncfusion<sup style="font-size:70%">&reg;</sup> assemblies used in the Build Server were from the Trial Installer, we should register the license key in the application for the corresponding version and platforms to avoid a trial license warning. | Yes | Use any developer trial license to [generate](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate) keys for Build Environments as well. |
| **Licensed Installer** | If the Syncfusion<sup style="font-size:70%">&reg;</sup> assemblies used in the Build Server were from the Licensed Installer, then there is no need to register the license keys.<br><br>You can [download](https://help.syncfusion.com/java-file-formats/installation/web-installer/how-to-download#download-the-license-version) and [install](https://help.syncfusion.com/java-file-formats/installation/web-installer/how-to-install) the licensed version of our installer. | No | Not applicable |

## See Also

* [How to Generate Syncfusion<sup style="font-size:70%">&reg;</sup> FileFormats License Key?](https://help.syncfusion.com/java-file-formats/licensing/how-to-generate)
* [How to Register Syncfusion<sup style="font-size:70%">&reg;</sup> License Key in FileFormats Application?](https://help.syncfusion.com/java-file-formats/licensing/how-to-register-in-an-application)