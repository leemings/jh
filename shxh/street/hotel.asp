<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then session.Abandon
umoney=rst("����")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>������ջ</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topmargin=100>
<p align=center><font color=0000FF size=6 face='��Բ'><%=Application("Ba_jxqy_systemname")%>������ջ</font></p>
<form action=sleep.asp method=post><table bgcolor=cccccc border=3 width=50% align=center><tr><td>�շѱ�׼��ÿ��ʮ������һ��Сʱ,��������<%=umoney%>����������Ϣ<font color=ff0000><%=umoney\60%></font>��Сʱ</td></tr><tr><td>��Ҫ��Ϣ<input type=text name=utime size=3 maxlength=3>��Сʱ</td></tr><tr align=center><td><input type=submit value=' ˯ �� '> <input type=button onclick='javascript:window.close();' value=' �� �� '></td></tr></table></form>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p></body>