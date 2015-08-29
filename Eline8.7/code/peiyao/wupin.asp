<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>药材♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="left">
<div align="center">现有药材一览<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div>
<div align="center"> <br>
    <table border="1" align="center" width="180" cellpadding="1" cellspacing="0" height="31">
      <tr align="center">
<td nowrap width="84" height="16">
<div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
</td>
<td nowrap width="49" height="16">
<div align="center"><font color="#FFFFFF" size="-1">数量 </font> </div>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
Response.Expires=0
rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
if not(isnull(rs("w6"))) then
data1=split(rs("w6"),";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
%>
<tr>
<td width="84" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=data3(0)%>
</font> </div>
</td>
<td width="49" height="3">
<div align="center"><font color="#FFFFFF" size="-1"><%=data3(1)%>
</font> </div>
</tr>
<%
erase data3
next 
erase data1
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</div>
</div>
</body>
</html>
