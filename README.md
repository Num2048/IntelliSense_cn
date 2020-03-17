Excel-DNA IntelliSense_cn
======================

Fork是因为要用这个，但是英语不太行，于是翻译一下文档，[点击跳转原文档](https://github.com/Excel-DNA/IntelliSense)。

Excel-DNA是一个将.NET与Excel集成的独立项目。 - 官网 http://excel-dna.net

结合Excel-DNA，你可以用C#、VB或者F#编写用于Excel的.xll插件，可以编写高性能的自定义公式(UDFs)、Ribbon选项卡等。

这个项目用于为Excel自定义公式编写提示，可以单独作为插件部署或者与Excel-Dna插件打包发布。

总览
--------
Excel默认不支持将用户定义的功能显示为工作表智能感知的一部分。 我们使用Windows和Excel的UI自动化支持来跟踪Excel界面的相关更改，并在适当时显示提示信息。

当前状态
--------------
该项目正在积极开发中，并准备进行初始测试。

带提示的Excel-Dna项目代码如下（C#）：
```C#
[ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
public static double AddThem(
	[ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] 
	double v1,
	[ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     
	double v2)
{
	return v1 + v2;
}
```
函数提示效果如下：

![Function Description](https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/FunctionDescription.PNG)

and when selecting the function, we get argument help

![Argument Help](https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/ArgumentHelp.PNG)


User-defined functions written in VBA (either in an add-in or regular Workbook) can also provide IntelliSense descriptions, either by embedding descriptions in the Workbook, or in an external file.

The first configuration being tested now, is where the IntelliSense display server is loaded as a separate add-in.

Getting Started
---------------

For existing Excel-DNA add-ins (v0.32 or later):
  * Download and load the latest ExcelDna.IntelliSense.xll or ExcelDna.IntelliSense64.xll from the [Releases](https://github.com/Excel-DNA/IntelliSense/releases) page.
  * IntelliSense should work automatically for functions that have descriptions in [ExcelFunction] and [ExcelArgument] attributes.

For VBA workbooks or add-ins:
  * Download and load the latest ExcelDna.IntelliSense.xll or ExcelDna.IntelliSense64.xll from the [Releases](https://github.com/Excel-DNA/IntelliSense/releases) page.
  * Either add a sheet with the IntelliSense function descriptions, or a separate xml file.

See the [Getting Started](https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started) page for more detail.

Future direction
----------------

Once a basic implementation is working, there is scope for quite a lot of enhancement. For example, we could add support for:

  * enum lists and other parameter selection and validation
  * links to forms or hyperlinks to help
  * enhanced argument selection controls, like a date selector

Support and participation
-------------------------
"We accept pull requests" ;-) 
Any help or feedback is greatly appreciated.

Please log bugs and feature suggestions on the GitHub 'Issues' page.

For general comments or discussion, use the Excel-DNA forum at https://groups.google.com/forum/#!forum/exceldna .

License
-------
This project is published under the standard MIT license.


  Govert van Drimmelen
  
  govert@icon.co.za
  
  18 June 2016
  
