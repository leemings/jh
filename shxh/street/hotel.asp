<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then session.Abandon
umoney=rst("银两")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>悦来客栈</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topmargin=100>
<p align=center><font color=0000FF size=6 face='幼圆'><%=Application("Ba_jxqy_systemname")%>悦来客栈</font></p>
<form action=sleep.asp method=post><table bgcolor=cccccc border=3 width=50% align=center><tr><td>收费标准：每六十两银子一个小时,你有现银<%=umoney%>两，可以体息<font color=ff0000><%=umoney\60%></font>个小时</td></tr><tr><td>我要体息<input type=text name=utime size=3 maxlength=3>个小时</td></tr><tr align=center><td><input type=submit value=' 睡 了 '> <input type=button onclick='javascript:window.close();' value=' 关 闭 '></td></tr></table></form>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p></body>