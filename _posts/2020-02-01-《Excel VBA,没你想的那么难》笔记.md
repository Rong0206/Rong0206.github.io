---
layout:     post                    
title:      "《Excel VBA,没你想的那么难》笔记"        
subtitle:   "Vba"
date:       2020-02-01             
author:     "CookBoy"                
header-img: img/post-bg-debug.png
catalog: true                       
tags:        学习
---
# Excel VBA,没你想的那么难
### 第一章 概述
介绍了VBA的作用：减少无脑地重复性Excel操作
1. 宏的录制：开发工具->录制宏（可设置快捷键）->使用相对引用->操作->停止录制
> 可以在 插入->按钮 里面设置按钮快捷键

2. 宏的源代码就是由VBA编写，可以在 宏->编辑 里面查看
> VB= Visual Basic For Applications

### 第二章 认识编辑工具
1. VBA编译器（Visual Basic Editor）打开方式：

- Excel窗口中输入Alt+F11
- 开发工具->Visual Basic
- 开发工具->查看代码
- 右键工作表->查看代码

2. 编译器窗口

![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/02EB434B8FE44800BC0C663CA8544F6D/8192)
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/77E7BC11D8BB4FF7B5C3C14139996E04/8172)

立即窗口可用于调试
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/E1D05D50B85748A0ACD3F84ED06ADBF3/8196)

3. VBA代码：Sub语序以"Sub 宏名"开头，以"End Sub"结束，如下图。
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/91511F9BA2CC4557BA300DE34EC6DB92/8212)

### 第三章VBA遵循规则
1. Excel中数据只有文本、数值、日期值、逻辑值、错误值5种类型
2. 
类型|类型名称|包含的数据及范围
---|---|---
布尔型|Boolean|逻辑值True或False
字节型|Byte|0到255的整数
整数型|Integer|
长整数型|Long|
单精度浮点型|Single|
双精度浮点型|Double|
货币型|Currency|
小数型|Decimal|
日期型|Date|
字符串型|String|
变体型|Variant|
对象型|Object|对象变量，用来引用对象
用户自定义类型|用户自定义|

3. 声明变量

方式|备注
---|---
Dim 变量名 As 数据类型|
Private 变量名 As 数据类型|私有变量
Public 变量名 As 数据类型|公有变量
Static 变量名 As 数据类型|声明为静态变量，程序结束后，静态变量会保持原值不变
4. 变量赋值
```
[Let] 变量名称 = 要储存的数据
#Let可省略
```
5. 对象赋值
用于储存工作簿、工作表、单元格等对象(Object)时用
```
Set 变量名称 = 要储存的对象名称

Dim sht As Worksheet  #定义一个工作表对象sht
Set sht = ActiveSheet #将活动工作表赋给变量sht
```
实例一：变量输入
```
Sub vary()
Dim temp
temp = 3000
Range("A1").Value = temp
End Sub
```
实例二：对象输入
```
Sub content()
Dim sht As Worksheet
Set sht = ActiveSheet
sht.Range("A1") = "lkr"
```
5. 声明变量方法
- 同时声明多个变量
```
Dim sht As Worksheet, IntCount As integer
```
- 使用类型声明符
```
Dim Str$
#即将Str声明为string
```
数据类型|类型声明字符
---|---
Integer|%
Long|&
Single|！
Double|#
Currency|@
String|$
- 声明变量可不指定类型
```
Dim Str
#将默认定义为Variant类型
```
若在代码前面加上"Option Explicit"，则强制所有代码声明变量
6. 不同作用域的变量
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/D24981CBD61E4EAA8BB3B8ED0FBC253D/8449)
7. 多个单变量组成的数组，将每个单变量称为数组的元素
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/ADEFC18BCED549C38110544E489F25E9/8461)
```
Dim arr(0 To 100) As Byte
```
通过索引可以查找数组数据
8. 声明多维数组
```
Dim arr(1 To 3, 1 To 5) As Byte
Dim arr(3, 5) As Byte
#默认起始索引号为0，可通过开头输入"OPTION BASE 0"设置
```
9. 声明动态数组
- 使用Public或Dim语句声明数组时，不能使用变量来确定数组的尺寸
```
Sub Test()
Dim a As Integer
a = Application.WorksheetFunction.CountA(Range("A:A"))
#VBA中使用工作表函数，需要借用Application对象的WorksheetFunction属性来调用
#错误如下：
Dim arr(1 To a) As String #错误！！！
#正确如下：
Dim arr() As String  #定义动态数组
ReDim arr(1 To a)  #ReDim可以重新定义数组范围，但不可以重新定义数组类型
End Sub
```
10. 使用Array创建数组
```
arr = array(1, 2, 3, 4, 5, 6)
```
11. 使用Split创建数组
```
arr = Split("一,二,三,四,五", ",")
#按照","拆分成数组
```
12. 通过单元格创建数组
```
Dim arr As Variant
arr = Range("A1,C3").Value
Range("E1,G3").Value = arr
```
13. 数组常用函数
- UBound求数组最大索引号
```
UBound(数组名称)
若为UBound(数组名称,1)，则表示第一维度的最大索引号
```
- LBound求数组最小索引号
```
LBound(数组名称)
```
- 求数组包含元素个数UBound(数组名称) - LBound(数组名称) + 1
- Join函数合并一维数组成字符串
```
txt = Join(arr,"@") #以@为分割符
```
14. 将数组保存的数据写入单元格
```
Range("A1").Value = arr(2)
Range("A1,A9").Value = Application.WorksheetFunction.Transpose(arr)
#将一维数组写入单元格时，单元格区域需为同一行，若要按列输入，可通过工作表中的transpose函数将数组转置为一列
```
15. 声明常量时要同时给常量赋值
```
Const 常量名称 As 数据类型 = 值
```
注意：常量也有不同的作用域，可用Public定义为公共常量
16. 引用对象
```
Application.Workbooks("Book1").Worksheets("Sheet2").Range("A2")
#若Book1是活动工作簿，可写成
Worksheets("Sheet2").Range("A2")
#若Sheet2是活动工作表，可写成
Range("A2")
```
17. 对象-属性-方法
- 对象和属性是相对的
- 方法就是在对象上执行的某个动作或操作
- 属性返回对象包含的内容或具有的特点，方法是对对象的一种操作
- 按ctrl+J可以调出方法/属性列表
18. VBA的四类运算符：算术运算符、比较运算符、文本运算符和逻辑运算符
19. 算术运算符

运算符|作用
--|--
+|求和
-|求差或求相反数
*|求积
/|求商
\|两数相除后所得商的整数
^|求一个数的某次方
Mod|两数相除后所得的余数
20.比较运算符

运算符|作用|返回结果
---|---|---
=|比较两个数据是否相等|相等返回True，否则返回False
<>|比较两个数据是否不相等|不等于返回True，否则翻回False
<|小于|
>|大于|
<=|小于或等于|
>=|大于或等于|
Is|比较两个对象的引用变量|若引用对象相同返回True，否则返回False
Like|比较两个字符串是否匹配|匹配时返回True，否则返回False

- 通配符

符号|作用|举例
--|--|--
*|代替任意多个字符|"李狗剩"Like"李狗*"=True
?|代替任意单个字符|"李狗剩"Like"李??"=True
#|代替任意单个数字|"李狗1"Like"李狗#"=True
[charlist]|代替charlist中任意一个字符|"I"Like""[A-Z]"=True
[!charlist]|代替不在charlist中任意一个字符|"I"Like""[!H-J]"=False
21. 文本运算符

文本运算符有+和&两种，他们都可以使得运算符左右两边的字符串合并为一个字符串
22.逻辑运算符
运算符|作用
---|---
And|与
Or|或
Not|非
Xor|异或
Eqv|等价
Imp|蕴含
22. 运算优先级
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/2C217B42ACC5459086FA41EB868A7766/8879)
23. IF语序
```
If ... Then 
... 
Else
...
End If
#例如
if range("A1").value >=60 then range("A2").Value = "及格" else range("A2").Value = "不及格"
```
多重if语序
```
#方法一：
If....Then
...
Else
    If...Then
    ...
    Else
    ...
    End If
End If
方法二：
If...Then
...
ElseIf...Then
...
ElseIf...Then
...
Else
...
End If
```
24. Select Case语句
```
#同样应用于多选问题
Select Case 表达式
    Case Is 条件表达式1
        ...
    Case Is 条件表达式2
        ...
    Case Is 条件表达式3
        ...
    Case Else
        ...
End Select
```
25. 循环语序
- For...Next语序
```
For <循环变量> = <初值> To <终值> [Step步长值]
    <循环体>
    [Exit For]
    <循环体>
Next [循环变量]
```
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/8E398DFA1333424A98350273731B94B5/8962)
- For Each...Next语序
```
#For Each...Next用于遍历集合或数组中的每个元素
For Each 变量 In 集合名称或数组名称
    语句块1
    [Exit For]
    语句块2
Next [元素变量]
```
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/0DD7D653B9C94CEA830AF68094B097DA/8999)
- Do While语序
```
#开头判别式
Do [While 循环条件]
    <循环体>
    [Exit Do]
    [循环体]
Loop
#结尾判别式
Do
    <循环体>
    [Exit Do]
    [循环体]
Loop [While 循环条件]
```
Do While中的Exit Do语句应该有条件判别，例子如下
```
If ... Then Exit Do
```
- Do Until语序
```
#开头判别式
Do [Until 循环条件]
    <循环体>
    [Exit For]
    <循环体>
Loop
#结尾判别式
Do
    <循环体>
    [Exit For]
    <循环体>
Loop [Until 循环条件]
```
26. GoTo语句
```
#在目标语句前加上一个带冒号的文本字符串或不带冒号的数字标签
x: mysum = mysum + i
If i <=100 Then GoTo x
```
27. With语句，简写代码
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/419F859405A943A2B0FCDAD44672144F/9076)
28. Sub语序
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/1373A40E3D5C4B41A9E7653BC06F201E/9083)
Public公共过程域的Public可以省略，默认为公共过程；宏对话框中只显示公共过程，私有过程没有显示，私有过程只有当前模块中可以调动
29. 在过程中调用过程
```
#方法1
过程名,参数1,参数2,...
#方法2
Call 过程名 (参数1,参数2,...)
#方法3
Application.Run "过程名",参数1，参数2,...
```
30. 向过程传递参数
```
Sub ShtAdd(shtcount As Integer)
#若在括号中加上ByVal，则引用过程中不会再改变参数的值，即shtcount=8不会被调用
    Worksheets.Add Count = shtcount
    shtcount = 8
End Sub
Call ShtAdd(2)
```
31. 函数
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/679897B6714348CDBAC3AC6655DCFDFD/9125)
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/C8142ED7D3774B839DADD1D982017654/9149)
32. 设置单元格颜色
```
Range("A1").Interior.Color = RGB(255,255,0)
```
33. 计算单元格颜色函数(可作为函数写法参考)
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/5CE7A7A464E544ED964638AE91A1029D/9146)
34. 易失性函数
```
#在代码中加入
Application.Volatile True
#每次数据更改，函数都会重新计算
```
35. 代码书写
- 用空格+下划线可以实现分行书写，即 _
- 用冒号可以实现多行书写到同一行，即 : 
- 单引号为注释，即 '

### 第四章 操作对象，解决工作中的实际问题
1. VBA中常用的对象

对象|对象说明
---|---
Application|代表Excel应用程序（如果在Word中使用VBA，就代表Word应用程序）
Workbook|代表Excel工作薄，一个Workbook对象代表一个工作薄文件
Worksheet|代表Excel工作表，一个Worksheet对象代表工作薄中一张普通的工作表
Range|代表Excel单元格，可以是单个单元格，也可以时单元格区域

2. 程序运算步骤数据更新关闭
```
Application.ScreenUpdating = False
```
3. 不再显示警告对话框
> 执行某写删除操作时会出行警告弹框
```
Application.DisplayAlerts = False
Application.DisplayAlerts = True
```
4. 调用Excel中的函数
```
Application.WorksheetFunction.XXXXXX
#注意：并不是所有工作表函数都能通过Worksheet调用
```
5. Application对象的常用属性
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/7F7DB38EDE7E42EBA64FB96BE20CA814/9159)
6. Workbook代表一个工作簿，workbooks代表当前打开的所有工作薄，即工作薄集合
7. 获取工作薄信息
```
ThisWorkbook.Name #获得工作簿名称
ThisWorkbook.Path #获得工作簿文件所在路径
ThisWorkbook.FullName #获得带路径的工作薄名称
```
8. 创建空白工作簿
```
#直接创建
Workbooks.Add
#指定模板
Workbooks.Add Template:= "D:\模板.xlsm"
#指定工作簿类型
Workbooks.Add Template:=xlWBATExcel4MacroSheet
```
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/3B0B3382CD4348BC9DA952FFBBC81786/9168)
9. 用Open打开工作簿
```
Workbooks.Open "D:\我的文件\模板.xlsm"
```
10. 用Activate激活工作簿
```
Workbooks("工作簿1").Activate
```
11. 保存文件
- 保存在原文件中
```
ThisWorkbook.Save
```
- 另存为
```
ThisWorkbook.SaveAs Filename:="D:\test.Xlsm"
```
- 另存为且不关闭原文件
```
ThisWorkbook,SaveCopyAs Filename:="D:\test.Xlsm"
```
12. 关闭工作簿
- 关闭所有工作簿
```
Workbooks.Close
```
- 关闭单个工作簿
```
Workbooks("Book1").Close
```
- 关闭并保存工作簿
```
Workbooks("Book1").Close True
```
13. ThisWorkbook与ActiveWorkbook
> ThisWorkbook是对代码所在工作薄的引用，ActiveWorkbook是对活动工作簿的引用
14. 引用工作表
```
#同理Worksheet代表一张工作表，Worksheets代表多张工作表的集合
Worksheets("sheet1")
```
15. 用Add新建工作表
- 活动工作表前插入一张工作表
```
Worksheets.Add
```
- 用before或after参数指定插入工作表位置
```
Worksheets.Add before/after:= Worksheets("sheet1")
```
- Count指定插入数量
```
Worksheets.Add Count:=3
```
16. 修改工作表名称
```
Worksheets(2).Name = "工作表"
ActiveSheet.Name = "工作表"
```
17. 删除工作表
```
Worksheets("Sheet1").Delete
```
18. 激活工作表
```
Worksheets(1).Activate
Worksheets(1).select
```
19. 将工作表复制/移动到指定位置
```
Worksheets(3).Copy before :=Worksheet(1)
Worksheets(3).Copy after :=Worksheet(3)
#复制工作簿中的第一张工作表到新工作簿中
Worksheet(1).Copy
#移动
Worksheets(3).Move before :=Worksheet(1)
Worksheets(3).Move after :=Worksheet(3)
#移动工作簿中的第一张工作表到新工作簿中
Worksheet(1).Move
```
20. 设置Visible属性
```
Worksheets(1).Visible = False  #隐藏
Wokrsheets(1).Visible = True   #显示
```
21. Worksheets和Sheets的区别
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/F88BFEFBDDEA4C6599DDE7CFC97AC84B/9297)
22. Range的引用
```
#引用多个不连续的单元格区域
Range("A1:A10,A4:E6,C3:D9").Select #用逗号分隔
#引用多个区域的公共区域
Range("B1:B10 A4:D6") #用空格分隔
#引用两个区域围成的矩形区域
Range("B6:B10","D2:D8").Select #双引号标注和逗号分隔
```
23. Cell引用单元格
```
#工作表对象.Cells(行号,列标)
ActiveSheet.Cells(3,4).Value = 20
```
24. 直接引用单元格（不能在括号中使用变量）
```
[B2]
[A1:D10]
[A1:A10,C1:C10,E1:E10] #三个单元格区域的并集
[B1:B10 A5:D5] #两个单元格区域的交集
[n] #被定义为n的单元格区域
```
25. 整行引用
```
Rows("3:10").Rows("1:1").Select
#选中第3行到第10行中的第1行
#同理整列采用Columns
```
26. 使用Union选择多个区域单元格
```
ThisWorkbook.Sheets(2).Application.Union(Range("A1:A10"), Range("D1:D5")).Select
```
27. offset参数
```
Range("B2,C3").Offset(5,3).Value=500
#表示向下移动5个单元格后向右移动3个单元格
```
28. Resize
```
Range("B2").Resize(5,4).Select
#重新扩展为5行4列
```
29. UsedRange
```
ActiveSheet.UsedRange.Select
#选中已经使用的所有单元格
```
30. CurrentRegion
```
Range("B5").CurrentRegion.Select
#返回指定单元格内的一个连续的矩形区域，遇到空行会阻断
```
31. Range的End属性
相当于返回该方向的最后一个非空单元格
```
Range("C5").End(xlUp)
```
可设置参数|参数说明
---|---
xlToLeft|相当于End+左方向键
xlToRight|相当于End+右方向键
xlUp|相当于End+上方向键
xlDown|相当于End+下方向键
32. Count属性，获得区域包含的单元格个数
```
ActiveSheet.UsedRange.Rows.Count
ActiveSheet.UsedRange.Columns.Count
```
33. Address可以获得单元格地址
34. 用Activate与Select都可以选中激活单元格
35. Copy复制单元格【Cut同理】
```
源单元格区域.Copy Destination:=目标单元格
#Destination:=可以省略
#注意：无论源单元格是区域有多大，目标单元格都可以只指定最左上单元格
```
36. Delete删除单元格
```
Range("B5").Delete Shift:=xlToLeft
#删除单元格后右侧单元格左移动，同理可以用xlUp等
Range("B5").EntireColumns.Delete
#删除整列，同理可用EntireRow
```
### 第五章 执行程序的自动开关---对象
1. 当某个事间发生后（如打开工作簿）自动运行的过程，我们称其为“事件过程”，事件过程也是Sub过程。
> 与普通的Sub过程不同，事件过程的作用域、过程名称及参数都不需要设置，也不能胡乱修改，其命名规则如下：
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/D48C6598415249899406362544028699/9620)

2. 常用事件：Open、Activate、Change
> SheetChange事件需要参数(ByVal Sh As Object, ByVal Target As Range),sh代表的是被修改的单元格所在的工作表，Target代表单元格，注意：SheetChange会令每一张工作表都应用
> 利用Application.EnableEvents = False来禁用事件，防止Change事件不停循环
> SelectionChange可以返回选择中的单元格位置

常用的WorkSheet事件，如下：
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/2B99196759D94C90A5D9313A9517AA3A/9648)

3. 常用的Workbook事件（Workbook事件会应用到所有Worksheet中）
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/8E1ADA5FB5CF4F43BAE8D1853E7C821C/9668)
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/CB722D40C94B449AA5035BDF4970A5AA/9670)

4. Application.onkey可以设置当键盘按下指定键/组合键时自动执行指定程序，但录制宏的方法更为便捷，故不提倡。
```
Application.onkey "+e","Hello"
#当按下shift+e时，执行Hello过程
```
5. Application.OnTime可以在指定的时间，执行指定的过程
```
Application.OnTime TimeValue("12:00:00"),"TellMe"
#中午12点时，执行TellMe过程
```

### 第六章 设计自定义的操作界面
1.表单按钮和ActiveX控件
> 表单控件的用法比较单一，只能在工作表中通过设置控件的格式或指定宏来使用，而ActiveX控件拥有很多属性和事件，不但可以在工作表中使用，还可以在用户窗体中使用
2. InputBox函数与InputBox方法的异同
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/A5F33102DF4649DA911B1ED702674B34/9764)
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/75F5F9559DBD439B8813372CBB970000/9770)
3. MsgBox类型
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/FCEB97707ADD434CB61A49DDBBA639ED/9775)
MsgBox返回值如下
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/497280CC359F4FA4B3218C5C65CA7103/9783)
4. 几个巧用的函数：

名称|用途
---|---
Application.FindFile|显示【打开】对话框
Application.GetOpenFilename|显示【打开】对话框并获取文件名
Application.GetSaveAsFilename|显示【另存为】对话框
Application.FileDialog|获取目录名称

5. 窗体
```
#窗体加载方法
Sub ShowForm()
    Load InputForm
    InputForm.Show
End Sub
#无模式窗体允许进行窗体外的其他操作
InpuForm.Show vbModeless
#默认为模式窗体
```
Initialze事件可以初始化窗体，如下为窗体的复合框选项中可选项目的设置
![](https://note.youdao.com/yws/public/resource/f4db12a79b98cc28e233f8e362e9afea/xmlnote/4C0826CF381748DEBB9D03312613F271/9851)
6. 本章主要讲如何设计一个窗体，该功能可以简易制作一个UI，然后编辑各个UI的代码，简化使用，同时对不同使用者友善，详细可翻看本书6.7章节

### 第七章 调试与优化编写的代码
1. 常见错误：编译错误、运行时错误、逻辑错误、
2. vba的三种模式：设计模式、运行模式、中断模式
3. 按F9设置断点，可以让程序运行到断点时暂停，再按F8，逐行运行
4. 用Stop也可设置断点
5. 可用Debug.Print将值输出到立即窗口检查
```
Debug.Print "i= " & i
```
6. 本地窗口可以看到所有变量的值以及数据类型
7. 监视窗口可以添加监视对象，实时查看数据
8. On Error 三种形式，通常On Error要放在程序开始处，要在错误发生前
```
On Error GoTo a #如果发生错误，则转到标签a的语句继续执行
On Error Resume Next  #忽略错误的代码，继续运行程序
On Error GoTo 0  #关闭错误捕捉
```
9. 如何让程序合理化(占用内存更少，运行更快)
- 声明变量为合适的数据类型
- 尽量避免使用Variant类型的变量
- 不要让变量一直保持在内存中，记得释放
```
Dim rng As Range
Set rng = Nothing
```
10. 将一维数组写入一列单元格时，应该将一维数组从行转置为列：Transpose()