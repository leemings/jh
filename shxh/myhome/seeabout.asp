<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ����,�Ա�,����,���,��ż,����,����,�ȼ�,����,����,����,����,����,����,����,��Ա,��Աʱ��,protect from �û� where ����='"&username&"'",conn
for i=0 to 8
	msg1=msg1&"<td>"&rst.Fields(i).Value&"</td>"
	msg2=msg2&"<td>"&rst.Fields(i+9).Value&"</td>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<LINK href="../chatroom/css.css" rel=stylesheet>
<div align=center>
<font color=0000FF>����״̬</font><hr>
  <table width=95% bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF border=3>
    <tr bgcolor=ffff00 align=center><td>����</td><td>�Ա�</td><td>����</td><td>���</td><td>��ż</td><td>����</td><td>����</td><td>�ȼ�</td><td>����</td></tr>
    <tr><%=msg1%></tr>
    <tr bgcolor=ffff00 align=center><td>����</td><td>����</td><td>����</td><td>����</td><td>����</td><td>����</td><td>��Ա</td><td>��Աʱ��</td><td>protect</td></tr>
    <tr><%=msg2%></tr>
  </table>
 <input type=button value='����' onclick=javascript:location.href='myhome.asp'; id=button1 name=button1>
</div>
</body>
