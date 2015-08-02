<%
username=session("Ba_jxqy_username")
usersex=session("Ba_jxqy_usersex")
if username="" then Response.Redirect "../error.asp?id=016"
if usersex="男" then 
	matesex="女"
else
	matesex="男"
end if
quest=server.HTMLEncode(Trim(Request.Form("quest")))
content=server.HTMLEncode(Trim(Request.Form("content")))
opt=server.HTMLEncode(Trim(Request.Form("opt")))
if content="" then Response.Redirect "../error.asp?id=043"
if instr(quest,"'")<>0  then Response.Redirect "../error.asp?id=900"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open "select 配偶 from 用户 where 姓名='"&username&"'",conn
mate=rst("配偶")
if opt="求婚" then
	if mate<>"无" then Response.Redirect "../error.asp?id=038"
	rst.Close
	rst.Open "select 对象 from 媒婆 where 申请人='"&username&"'",conn
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
else
	if mate="无" then Response.Redirect "../error.asp?ID=042"
	sqlstr="select * from 媒婆 where 申请人='"&username&"'"
	rst.Close
	rst.Open sqlstr,conn
	if not(rst.EOF or rst.BOF) then Response.Redirect "../error.asp?id=042"
	sqlstr="insert into 媒婆(申请人,对象,说明,时间,操作) values('"&username&"','"&mate&"','"&content&"',"&nowtimetype&",True)"
end if
rst.Close
set rst=nothing
conn.Execute sqlstr
conn.Close 
set conn=nothing
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3秒钟后自动返回<br>
  <font color=FF0000><img src="../images/error.gif" width="32" height="32"> 操作请求完成，谢谢您的光顾</font></div>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>