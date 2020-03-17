Excel-DNA IntelliSense_cn
======================

Fork是因为要用这个，但是英语不太行，于是翻译一下文档，[点击跳转原文档](https://github.com/Excel-DNA/IntelliSense)。

Excel-DNA是一个将.NET与Excel集成的独立项目。 - 官网 http://excel-dna.net

结合Excel-DNA，你可以用C#、VB或者F#编写用于Excel的.xll插件，可以编写高性能的自定义公式(UDFs)、Ribbon选项卡等。

这个项目用于为Excel自定义公式编写提示，可以单独作为插件部署或者与Excel-Dna插件打包发布。

总览
--------
Excel默认不支持将用户定义的功能显示为工作表自动提示的一部分。 我们使用Windows和Excel的UI自动化支持来跟踪Excel界面的相关更改，并在适当时显示提示信息。

当前状态
--------------
该项目正在积极开发中，并准备进行初始测试。

带提示的Excel-Dna项目代码如下（C#）：
```C#
[ExcelFunction(Description = "一个测试函数，返回两个数的和")]
public static double AddThem(
	[ExcelArgument(Name = "数值1", Description = "这是被计算的第一个数字")] 
	double v1,
	[ExcelArgument(Name = "数值2", Description = "这是被计算的第二个数字")]     
	double v2)
{
	return v1 + v2;
}
```
函数描述效果如下：

![Function Description](https://raw.github.com/num2048/IntelliSense/master/Screenshots/FunctionDescription.PNG)

输入参数时，可以看到提示

![Argument Help](https://raw.github.com/num2048/IntelliSense/master/Screenshots/ArgumentHelp.PNG)


用户在自定义VBA函数时 (代码在插件或者工作簿中) ，把相应代码写在工作簿或者外部文件中，也可以使用自动描述/提示。

目前的测试是基于把IntelliSense作为单独加载项进行的.

开始使用
---------------

使用Excel-DNA插件开发时 (0.32之后的版本):
  * 从[Releases](https://github.com/Excel-DNA/IntelliSense/releases)下载并加载最新版ExcelDna.IntelliSense.xll或ExcelDna.IntelliSense64.xll（64位Excel）。
  * IntelliSense将会自动识别自定义函数代码中的[ExcelFunction]和[ExcelArgument]属性。

对于VBA工作簿或者插件:
  * 从[Releases](https://github.com/Excel-DNA/IntelliSense/releases)下载并加载最新版ExcelDna.IntelliSense.xll或ExcelDna.IntelliSense64.xll（64位Excel）。
  * 添加一个包含IntelliSense函数描述的工作表，或者添加一个独立的xml文件。
  
查看 [开始使用](https://github.com/num2048/IntelliSense/wiki/快速入门) 页面获取更多信息.

未来目标
----------------

基础部分已经可以使用，但是还有可以优化的地方. 例如我们需要优化以下方面:

  * 枚举列表和其他参数的选择及验证
  * 添加帮助窗体或超链接
  * 更强大的参数选择控件，例如日期选择器

参与和支持
-------------------------
"我们接受Pull requests" ;-) 

我们感谢任何帮助或反馈。

请在GitHub[“Issues”](https://github.com/Excel-DNA/IntelliSense/issues)页面上记录错误和功能建议。

对于一般的评论或讨论，请使用https://groups.google.com/forum/#!forum/exceldna 上的Excel-DNA论坛。

开源协议
-------
项目基于MIT license开源.


  Govert van Drimmelen
  
  govert@icon.co.za
  
  18 June 2016
  
