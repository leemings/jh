<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Db_path="../DataBackup/" & Request.form("getmdb")
Db_path=server.mappath(Db_path) 
getname=replace(Request.form("getname"),"'","")
if getname=Application("aqjh_user") then
 response.write "不能对站长操作！"
response.end
end if
on error resume next
sql="select * from [" & Db_path & "].[用户] where 姓名='" & getname & "'"
rs.open sql,conn,1,1
if rs.recordcount=0 then
 response.write "找不到用户！"
response.end
end if
sql="delete * from [用户] where 姓名='" & getname & "'"
conn.Execute sql
sql2="Insert INTO [用户] select * from [" & Db_path & "].[用户] where 姓名='" & getname & "'"
conn.Execute sql2
conn.close
set conn=nothing
if err.number<>0 then 
response.write "SQL语法错误，恢复账号失败！"
response.end
end if
%>
<html>
<head>
<title>账号急救</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center">**************<b></b><b>账号急救</b>**************<b></b></p>
<p><font size="2" color="#FF0000">您已经成功恢复了<%=getname%>的账号<br></font><br></p>