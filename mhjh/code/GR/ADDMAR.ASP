<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")                                  
conn.Open Application("yx8_mhjh_connstr")                                  
set rs=server.CreateObject("adodb.recordset")                                  
rs.Open "select 性别 from 用户 where 姓名='"&username&"'",conn
usersex=rs("性别") 
if usersex="男" then
matesex="女"
else
matesex="男"
end if
rs.Close                                  
set rs=nothing  
quest=server.HTMLEncode(Trim(Request.Form("quest")))
content=server.HTMLEncode(Trim(Request.Form("content")))
opt=server.HTMLEncode(Trim(Request.Form("opt")))
if content="" then Response.Redirect "../error.asp?id=043"
if instr(quest,"'")<>0  then Response.Redirect "../error.asp?id=900"
if len(quest)>12 or len(content)>50 then Response.Redirect "../error.asp?id=530"
if instr(quest,"file:")<>0 or instr(quest,"script:")<>0 or instr(quest,"'")<>0 or instr(quest,chr(34))<>0 or instr(quest,"\")<>0 then Response.Redirect "../error.asp?id=524"
if instr(content,"file:")<>0 or instr(content,"script:")<>0 or instr(content,"'")<>0 or instr(content,chr(34))<>0 or instr(content,"\")<>0 then Response.Redirect "../error.asp?id=524"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open "select 配偶 from 用户 where 姓名='"&username&"'",conn
mate=rst("配偶")
if opt="求婚" then
if mate<>"无" then Response.Redirect "../error.asp?id=038"
rst.Close
sql="select 对象 from 媒婆 where 申请人='"&username&"'"
 Set Rst=connt.Execute(sql)
if not (rst.EOF or rst.BOF)then
if rst("对象")<>quest then
Response.Redirect "../error.asp?ID=039"
else
Response.Redirect "../error.asp?id=040"
end if
end if
rst.Close
rst.Open "select * from 用户 where 姓名='"&quest&"' and 性别='"&matesex&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=041"
sqlstr="insert into 媒婆(申请人,对象,说明,时间,操作) values('"&username&"','"&quest&"','"&content&"',"&nowtimetype&",False)"
 Set Rst=connt.Execute(sqlstr)
else
if mate="无" then Response.Redirect "../error.asp?ID=042"
sqlstr="select * from 媒婆 where 申请人='"&username&"'"
 Set Rst=connt.Execute(sqlstr)
if not(rst.EOF or rst.BOF) then Response.Redirect "../error.asp?id=042"
sqlstr="insert into 媒婆(申请人,对象,说明,时间,操作) values('"&username&"','"&mate&"','"&content&"',"&nowtimetype&",True)"
 Set Rst=connt.Execute(sqlstr)
end if
conn.Close
set conn=nothing
connt.Close
set connt=nothing
%>
<head><title><%=Application("yx8_mhjh_systemname")%></title><LINK href="../style.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center><font color="#FFFFFF">3秒钟后自动返回<br>
  &nbsp;操作请求完成，谢谢您的光顾</font></div> 
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;"><font color="#FFFFFF">返回</font></a></p> 
</body> 
