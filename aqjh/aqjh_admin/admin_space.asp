<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
%>
<head>
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<table cellpadding=0 cellspacing=0 border=0 width=95% align=center>
<tr>
<td>
<table cellpadding=3 cellspacing=1 border=0 width=100%>
<tr>
<td align=center colspan="2">欢迎<b><%=aqjh_name%></b>进入管理页面
</td>
</tr>
<tr>
<td width="20%" valign=top>
</td>
<td width="80%" valign=top>
<% 
dim b_d3
Function ShowSpaceInfo(drvpath)
dim fso,d,size,showsize
showsize=0
set fso=server.createobject("scripting.filesystemobject") 
drvpath=server.mappath(drvpath) 
set d=fso.getfolder(drvpath) 
size=d.size
showsize=size 
b_d3= "&nbsp;Byte"
if size>1024 then
	size=(size\1024)
	b_d3="&nbsp;KB"
	showsize=size 
end if
if size>1024 then
	size=(size/1024)
	b_d3="&nbsp;MB" 
	showsize=formatnumber(size,2)
end if
if size>1024 then
	size=(size/1024)
	b_d3="&nbsp;MB" 
	showsize=formatnumber(size,2) 
end if 
ShowSpaceInfo=showsize
Set d=Nothing
Set Fso=Nothing
End Function 

Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../chat")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath) 
if method="All" then 
	size=d.size
elseif method="Program" then
	set fc=d.Files
	for each f1 in fc
	size=size+f1.size
	next 
end if 
	showsize=size & "&nbsp;Byte" 
if size>1024 then
	size=(size\1024)
	showsize=size & "&nbsp;KB"
end if
if size>1024 then
	size=(size/1024)
	showsize=formatnumber(size,2) & "&nbsp;MB" 
end if
if size>1024 then
	size=(size/1024)
	showsize=formatnumber(size,2) & "&nbsp;GB" 
end if 
	response.write "<font face=verdana>" & showsize & "</font>"
	Set d=Nothing
	Set Fso=Nothing
end sub 

Function Drawbar(drvpath)
dim fso,drvpathroot,d,size,totalsize,barsize
set fso=server.createobject("scripting.filesystemobject")

drvpath=server.mappath(drvpath) 
set d=fso.getfolder(drvpath)
size=d.size
if size>1024 then size=(size\1024)
if size>1024 then size=(size/1024)
if size>1024 then size=(size/1024)
Drawbar=size * 3
Set d=Nothing
Set Fso=Nothing
End Function 

Function Drawspecialbar()
dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
set fso=server.createobject("scripting.filesystemobject")
drvpathroot=server.mappath("../chat")
drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
set d=fso.getfolder(drvpathroot)
size=d.size
if size>1024 then size=(size\1024)
if size>1024 then size=(size/1024)
if size>1024 then size=(size/1024)
Drawspecialbar=size * 3
Set d=Nothing
Set Fso=Nothing
End Function 

%>
<table width=550 cellspacing=1 cellpadding=0 > 
<tr>
<td height=25 >
&nbsp;&nbsp;
</td>
</tr> 
<tr>
<td> 
<blockquote> 
<%
fsoflag=1
if fsoflag=1 then
%>
<br> 
江湖数据占用空间：&nbsp;<img src="../images/bar1.gif" width=<%=drawbar("../aqjh_data")%> height=10>&nbsp;<font face=verdana><%=showSpaceinfo("../aqjh_data") & b_d3%></font><br><br>
江湖占用空间总计：<br><img src="../images/bar1.gif" width=<%=Drawspecialbar()%> height=10>&nbsp;<br><%showspecialspaceinfo("All")%>
<%
else
response.write "<br><li>本功能已经被关闭"
end if
%>
</blockquote> 
</td>
</tr>
</table> 
</td> 
</body>