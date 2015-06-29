AM Report开发使用说明书</P>
</P>
作者：Angus Meng</P>
联系方式：mqlxp@163.com</P>
</P>
目  录</P>
</P>
1	环境需求</P>
1.1	操作系统</P>
1.2	运行环境</P>
1.3	软件环境</P>
1.4	开发环境</P>
2	类库说明</P>
2.1	属性</P>
2.2	方法</P>
3	模板格式</P>
3.1	表格模式</P>
3.2	单元格模式</P>
4	配置软件使用说明</P>
4.1	数据源配置</P>
4.2	数据读取配置</P>
4.3	数据表格配置</P>
4.4	数据表格模式</P>
4.5	单元格模式</P>
4.6	配置文件</P>
4.7	更多功能</P>
</P>
1	环境需求
1.1	操作系统
Windows XP SP3, Windows 7, Windows 8。
1.2	运行环境
需要安装.Net Framework 4.0或以上版本。
1.3	软件环境
需要安装Excel 2007。
1.4	开发环境
Microsoft Visual Studio 2010或以上版本。

2	类库说明
2.1	属性
Mode ：模式设置，其值为 table 或 cell
DataSource ：数据源配置，机器名（或IP地址）[+ 实例名称]
InitialCatalog :初始数据库名称
PersistSecurityInfo：安全性设置，True 或 False
UserID：用户名
Password：密码
Row：表格模式中数据打印的起始行
Col：表格模式中数据打印的起始列
Table：需要查询的表或视图的名称
Top：返回结果集中的前N行
Where：查询条件语句
OrderBy：排序设置，默认为空，不进行排序
GroupBy：分组设置，默认为空，不进行分组
Items：（字符串数组）表格模式中查询列名设置
ItemCells：（字符串数组）单元格模式中查询列名设置
2.2	方法
BuildOptionsFromFile(ByVal FileName As String) As Boolean 
导入配置文件，FileName为.opt配置文件的完整路径名称。

Export(ByRef DataSheet As Microsoft.Office.Interop.Excel.Worksheet)
导出数据到Excel数据表，DataSheet为Excel数据表指针。
Export方法要和BuildOptionsFromFile方法配合使用，或者先手动配置相应的属性，然后再调用Export方法将数据打印到Excel数据表。

Export(ByRef DataSheet As Microsoft.Office.Interop.Excel.Worksheet, ByVal OptionFile As String)
此方法将导入配置文件和导出数据到Excel数据表合并在一起，Export(DataSheet,OptionFile)就是BuildOptionsFromFile(FileName)和Export(DataSheet)的合并。

Register(ByVal UserID As String, ByVal RegKey As String) As Boolean
授权导入，请将分发给你的用户名和注册串号分别赋值给UserID和RegKey调用此函数进行授权。

3	模板格式
模板文件是以.opt格式结尾的文件，其内容为文本文件。其中包含配置内容，根据配置内容不同又分为两种模式：表格模式和单元格模式。
3.1	表格模式
[mode]:表格模式设置为table
[data soure]:数据源名称或IP地址
[initial catalog]:初始数据库名称
[persist security info]:安全性设置，True 或 False
[user id]:SQL Server 用户名
[password]:密码
[Row]: 数据在Excel表格中打印的起始行
[Col]: 数据在Excel表格中打印的起始列
[table]:数据库中数据表或视图的名称
[top]:返回结果集中的前N行，此项为空则返回所有行
[where]:查询条件
[order by]:排序设置，默认为空，不进行排序
[group by]:分组设置，默认为空，不进行分组
[items]:查询的数据项（列名），此配置行必须放在所有配置项的最后。且此配置项后不放参数，需要查询的列名放到此配置项的下面

表格模式配置文件示例：
[mode]:table
[data soure]:(local)
[initial catalog]:gangcheng
[persist security info]:False
[user id]:sa
[password]:mql
[Row]:4
[Col]:2
[table]:xiaoshijilu
[top]:
[where]:date_str = '2015-03-01'
[order by]:
[group by]:
[items]:
JS_LL1+JS_LL2,R.0.3
null
js_ph,I.0.0
js_cod,D.ON.OFF
cs_ll1+CS_LL2
null
cs_ph
cs_cod

示例说明：
[where]:date_str = '2015-03-01'
返回左右所有date_str = '2015-03-01'的数据。
[items]:此配置项下面的参数后可以设置数据格式。R——实型数据，R.0.3表示该数据为实型数据，且保留3位小数。I——整型数据，I.0.0设置该数值以整型显示。D——布尔型数据，D.ON.OFF表示该数据为0时显示OFF，为1（或大于1）时显示ON。
注意：R.X.Y中，X恒为0；I.X.Y中X、Y恒为0。R、I、D格式设置与前面列名以英文逗号“,”分割。

3.2	单元格模式
[mode]:单元格模式设置为 cell
[data soure]:数据源名称或IP地址
[initial catalog]:初始数据库名称
[persist security info]:安全性设置，True 或 False
[user id]:SQL Server 用户名
[password]:密码
[table]: 数据库中数据表或视图的名称
[top]: 返回结果集中的前N行，此项为空则返回所有行
[where]: 查询条件
[order by]:排序设置，默认为空，不进行排序
[group by]:分组设置，默认为空，不进行分组
[itemcells]:查询的数据项（列名），此配置行必须放在所有配置项的最后。且此配置项后不放参数，需要查询的列名放到此配置项的下面,格式为“ColName,Row,Col,X.X.X”

单元格模式配置文件示例：
[mode]:cell
[data soure]:(local)
[initial catalog]:gangcheng
[persist security info]:False
[user id]:sa
[password]:mql
[table]:xiaoshijilu
[top]:
[where]:date_str = '2015-03-01'
[order by]:
[group by]:
[itemcells]:
avg(JS_LL1+JS_LL2),1,1,R.0.3
avg(js_ph),2,2,I.0.4
avg(js_cod),3,3,D.On.Off
avg(cs_ll1+CS_LL2),4,4
avg(cs_ph),5,5
avg(cs_cod),6,6

示例说明：
[itemcells]:下面的每一项必须设置单元格位置，即Row、Col的值。另外可以设置数据的格式，设置方法同表格模式中[items]项。

4	配置软件使用说明

4.1	数据源配置
数据源：数据库所在计算机的机器名或IP地址[+ 数据库实例名称]
用户名：数据库用户名
密码：用户名所对应的密码
系统权限：勾选此项将忽略用户名和密码，使用系统用户权限登录数据库
连接：点击该按钮，将会在数据库列表中显示该实例下所有数据库名称
依次点击>>按钮将分别在数据表、数据列中显示相应内容
4.2	数据读取配置
模式：分为数据表和单元格两种模式，对应配置文件中的[mode]:table和[mode]:cell
起始行、列：设置数据表模式下数据打印的其实行和列
前N行：设置返回查询结果的前N行数据
条件：查询条件
排序：查询结果的排序
分组：分组查询设置
4.3	数据表格配置
4.4	数据表格模式
数据表格模式时，只显示一列。将数据列列表中的内容用鼠标左键拖拽到下表中相应的单元格中，按显示顺序排列好。最终数据打印时将按照起始行、列设置的位置进行打印。
注：双击单元格可以删除该单元格中的内容。
4.5	单元格模式
单元格模式时，将显示全部行。将数据列列表中的内容用鼠标拖拽到下表中相应的单元格。如果数据查询结果只有一条，将按照设置好的位置打印数据。如果有多条数据，将行数据递增向下打印。
4.6	配置文件
配置完成后点击“保存”按钮保存配置文件。保存后单击“测试”按钮可以测试数据查询和打印是否符合设计要求。
单击…按钮可以选择配置文件，选择后单击“测试”按钮可以对指定的配置文件进行测试。
4.7	更多功能
如需更多功能请手动修改配置文件。
