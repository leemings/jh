<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 洞房花烛 from 用户 where 姓名='"&username&"'",conn
mate=rst("洞房花烛")
rst.Close
sql="select * from 洞房 where ID="&id&" and 对象='"&username&"'"
 Set Rst=connt.Execute(sql)
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
opt=rst("操作")
questname=rst("申请人")
if opt=False and mate="未" then
conn.Execute "update 用户 set 洞房花烛='"&questname&"' where 姓名='"&username&"'"
conn.Execute "update 用户 set 洞房花烛='"&username&"' where 姓名='"&questname&"'"
connt.Execute "delete * from 洞房 where 申请人='"&questname&"' and 对象='"&username&"' and 操作=False"
connt.Execute "delete * from 洞房 where 申请人='"&username&"' and 对象='"&questname&"' and 操作=False"
msg="春宵一夜值千金，好好享受这美好时光吧！"
elseif opt=True and mate<>"未" then
conn.Execute "update 用户 set 洞房花烛='未' where 姓名='"&username&"'"
conn.Execute "update 用户 set 洞房花烛='未' where 姓名='"&questname&"'"
connt.Execute "delete * from 洞房 where 申请人='"&questname&"' and 对象='"&username&"' and 操作=True"
connt.Execute "delete * from 洞房 where 申请人='"&username&"' and 对象='"&questname&"' and 操作=True"
msg="请求完成"
else
msg="哈哈，请求无效阿：）"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
connt.Close
set connt=nothing
%>
<head><title></title><LINK rel=stylesheet href='../../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif' topmargin=150>
<p align=center>3秒钟后自动返回<br> <font color=FF0000><%=msg%></font></p>
<p align=center><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>
