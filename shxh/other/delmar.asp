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
rst.Open "select * from ý�� where ID="&id&" and ������='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
rst.Close
set rst=nothing
conn.Execute "delete * from ý�� where ������='"&username&"' and id="&ID
conn.Close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='../chatroom/css.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=bgcolor%>" background="<%=bgimage%>" topmargin=150>
<p align=center>3���Ӻ��Զ�����<br> <font color=FF0000>�õ������������㳷������</font></p>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>