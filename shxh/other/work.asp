<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from �� where id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=103"
ucondition=rst("����")
uresult=rst("����")
umoney=rst("����")
rst.Close
rst.Open "select * from �û� where ����='"&username&"' and "&ucondition,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=047"
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "update �û� set "&uresult&" where ����='"&username&"' and "&ucondition
conn.CommitTrans
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>������"&umoney&"�����ӣ����պ��ˣ���ӭ������<br>�����Ӻ��Զ�����<br><a href='javascript:history.back();'>����</a></p></body>"
%>