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
rst.Open "select ��ż from �û� where ����='"&username&"'",conn
mate=rst("��ż")
rst.Close
rst.Open "select * from ý�� where ID="&id&" and ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=044"
opt=rst("����")
questname=rst("������")
if opt=False and mate="��" then
	conn.Execute "update �û� set ��ż='"&questname&"' where ����='"&username&"'"
	conn.Execute "update �û� set ��ż='"&username&"' where ����='"&questname&"'"
	conn.Execute "delete * from ý�� where ������='"&questname&"' and ����='"&username&"' and ����=False"
	conn.Execute "delete * from ý�� where ������='"&username&"' and ����='"&questname&"' and ����=False"
	msg="����úϣ�����˳�ģ�"
elseif opt=True and mate<>"��" then
	conn.Execute "update �û� set ��ż='��' where ����='"&username&"'"
	conn.Execute "update �û� set ��ż='��' where ����='"&questname&"'"
	conn.Execute "delete * from ý�� where ������='"&questname&"' and ����='"&username&"' and ����=True"
	conn.Execute "delete * from ý�� where ������='"&username&"' and ����='"&questname&"' and ����=True"
	msg="�������"
else
	msg="��ã�������Ч����"	
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='../chatroom/css.css'><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=bgcolor%>" background="<%=bgimage%>" topmargin=150>
<p align=center>3���Ӻ��Զ�����<br> <font color=FF0000><%=msg%></font></p>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>