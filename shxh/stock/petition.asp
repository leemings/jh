<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" then stock="������湤����"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>�깺ԭʼ��</h3><hr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ��� from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=016"
umoney=rst("���")
rst.Close
rst.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
unum=umoney\rst("���м�")
uallownum=unum
if uallownum>rst("ʣ��ɷ�") then uallownum=rst("ʣ��ɷ�")
msg=msg&stock&"��ԭʼ�ɷݻ�ʣ��"&rst("ʣ��ɷ�")&"��,��ӵ�����д��"&umoney&",<font color=FF0000>���깺"&uallownum&"��</font><form action='petitionnow.asp' method=post><table><tr><td>��Ʊ</td><td><input type=text value='"&stock&"' size=14 readonly name='stock'></td></tr><tr><td>����</td><td><input type=text value='"&uallownum&"' size=9 name='num'></td></tr><tr><td align=center colspan=2><input type=submit value='  �� �� '></td></tr></table></form>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>