<!-- #include file="conn.asp" -->
<!-- #include file="inc/function.asp" -->
<%
dim Title_Name,CategoryName,CategoryName_CHS,username,URL,kicktime

'栏目变量名
CategoryName="SoftDown"   '不要改动

'网站名称
Title_Name="一线网络 ==>> " '可以改动

'频道名称
CategoryName_CHS="E线江湖" '可以改动

UserName= session("sUserName") '不要改动

'运行环境
SystemType="Windows.Net,WinXP,Win2000,NT,WinME,Win9X,Dos,Linux,Unix,Mac,ASP环境,PHP环境,CGI环境"  '可以改动

'添加软件时默认运行环境
DefaultSystemType="WinXP,Win2000,NT,WinME,Win9X" '可以改动

'设置删除在线数据表中多少分钟内不活动用户,单位为分钟

kicktime=2400

URL="http://www.198962.com" '下载栏目网址 可以改动

FolderPath="../Down/"  '静态网页存放文件夹，相对于管理程序文件夹的位置。

%>