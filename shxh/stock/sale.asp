<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"' and ����>0",conn
if not (rst.EOF or rst.BOF) then  Response.Redirect "../error.asp?id=067"
rst.Close
rst.Open "select tc.��Ȩ,ts.�ּ� from �ֹ� tc inner join ��Ʊ ts ON  ts.��Ʊ����=tc.���� where tc.�ֹ���='"&username&"' and tc.����='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=051"
unum=rst("��Ȩ")
uprice=rst("�ּ�")
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
if uprice="" then Response.Redirect "../error.asp?id=100"
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>ʵʱ����</h3><hr><table bgcolor=CCCCCC><form action=salenow.asp method=post><table><tr><td bgcolor=FFFF00 align=center>���۹�Ʊ</td></tr><tr><td><input type=text size=14 value='"&stock&"' readonly name='stock'></tr><tr><td>����:1-"&unum&"��</td></tr><tr><td><input type=text value='"&unum&"' size=9 name='unum'></td></tr><tr><td align=center>�ۼ�</td></tr><tr><td><input type=text value='"&uprice&"' size=9 name='uprice'></td></tr><tr><td align=center><input type=submit value=' �� �� '> <input type=button value=' �� �� ' onclick='javascript:history.back();'></td></tr></table></form></body>"
%>