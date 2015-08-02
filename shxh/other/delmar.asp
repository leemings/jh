<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from 媒婆 where ID="&id&" and 申请人='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
rst.Close
set rst=nothing
conn.Execute "delete * from 媒婆 where 申请人='"&username&"' and id="&ID
conn.Close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='../chatroom/css.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=bgcolor%>" background="<%=bgimage%>" topmargin=150>
<p align=center>3秒钟后自动返回<br> <font color=FF0000>好的我们马上替你撤消请求</font></p>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>