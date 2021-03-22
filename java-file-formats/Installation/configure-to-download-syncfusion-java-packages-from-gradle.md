---
title: Download Syncfusion Java packages from Gradle | Syncfusion
description: This section demonstrates how to configure and download required Jars from Gradle (Jar configuration)
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
repositories&nbsp;{<br />
&nbsp;&nbsp;&nbsp;maven&nbsp; {<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:green;font-size:13px;font-style:italic">&nbsp;&nbsp;//Syncfusion maven repository to download the artifacts</span>.<br />
&nbsp;&nbsp;&nbsp;&nbsp;url "https://jars.syncfusion.com/repository/maven-public/"<br />
}<br />
}
</td>
</tr>
</table>

The following command shows how to refer to the Syncfusion package in Gradle, which needs to be used in your project as the dependency.

<table>
<tr>
<td>
	dependencies &nbsp;{<br />
 &nbsp;&nbsp;&nbsp;&nbsp;implementation 'com.syncfusion:syncfusion-docio:18.4.0.30'<br />
}
</td>
</tr>
</table>
