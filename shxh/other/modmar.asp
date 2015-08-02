<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 配偶 from 用户 where 姓名='"&username&"'",conn
mate=rst("配偶")
rst.Close
rst.Open "select * from 媒婆 where ID="&id&" and 对象='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
opt=rst("操作")
questname=rst("申请人")
if opt=False and mate="无" then
	conn.Execute "update 用户 set 配偶='"&questname&"' where 姓名='"&username&"'"
	conn.Execute "update 用户 set 配偶='"&username&"' where 姓名='"&questname&"'"
	conn.Execute "delete * from 媒婆 where 申请人='"&questname&"' and 对象='"&username&"' and 操作=False"
	conn.Execute "delete * from 媒婆 where 申请人='"&username&"' and 对象='"&questname&"' and 操作=False"
	msg="百年好合！事事顺心！"
elseif opt=True and mate<>"无" then
	conn.Execute "update 用户 set 配偶='无' where 姓名='"&username&"'"
	conn.Execute "update 用户 set 配偶='无' where 姓名='"&questname&"'"
	conn.Execute "delete * from 媒婆 where 申请人='"&questname&"' and 对象='"&username&"' and 操作=True"
	conn.Execute "delete * from 媒婆 where 申请人='"&username&"' and 对象='"&questname&"' and 操作=True"
	msg="请求完成"
else
	msg="你好，请求无效：）"	
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='../chatroom/css.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=bgcolor%>" background="<%=bgimage%>" topmargin=150>
<p align=center>3秒钟后自动返回<br> <font color=FF0000><%=msg%></font></p>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>