<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 车 where id="&id,conn
tp=rs("图片")
rs.close
conn.execute "delete * from 车 where id="&id
conn.execute "update 用户 set 车='无' where 车='" & tp & "'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>无标题文档</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../JHIMG/bk_Hc3w.gif">
<p align="center">删除交通工具成功</p>
<p align="center"><a href="manjtgj.asp">返回</a> </p>
</body>
</html>
