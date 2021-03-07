---
title: Configure to download Syncfusion Java packages from Gradle | Syncfusion
description: This section illustrate how to Download JAR from Gradle
platform: java-file-formats
control: general
documentation: UG
---
# Configure to download Syncfusion Java packages from Gradle

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
implementation 'com.syncfusion:syncfusion-docio:18.4.0.30'
}
</td>
</tr>
</table>
