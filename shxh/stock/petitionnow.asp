<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.form("stock")
num=Request.Form("num")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(num) then Response.Redirect "../error.asp?id=024"
num=clng(num)
if num<1 then num=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>�깺ԭʼ��</h3><hr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ��� from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=016"
umoney=rst("���")
rst.Close
rst.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
uprice=rst("���м�")
unum=umoney\uprice
uallownum=rst("ʣ��ɷ�")
if uallownum<unum then uallownum=unum
if num>unum then Response.Redirect "../error.asp?id=030"
if num>uallownum then num=uallownum
umoney=num*uprice
rst.Close
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"'",conn
conn.BeginTrans
if rst.EOF or rst.BOF then
	conn.Execute "insert into �ֹ�(�ֹ���,ʱ��,����,��Ȩ,����,���,����) values('"&username&"',now(),'"&stock&"',"&num&",0,"&uprice&",0)"
else
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&num&" where �ֹ���='"&username&"' and ����='"&stock&"'"
end if
conn.Execute "update ��Ʊ set ʣ��ɷ�=ʣ��ɷ�-"&num&" where ��Ʊ����='"&stock&"'"
conn.Execute "update �û� set ���=���-"&umoney&" where ����='"&username&"'"
rst.Close
set rst=nothing
if conn.Errors.Count>0 then
	msg=msg&"<p align=center>����<font color=FF0000>ʧ��</font>,�����Ӻ��Զ�����<br><a href='javascript:history.back();'>����</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
	conn.RollbackTrans
else
	msg=msg&"<p align=center>����<font color=0000FF>���</font>,�����Ӻ��Զ�����<br><a href='javascript:history.back();'>����</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
	conn.CommitTrans
end if	
conn.Close
set conn=nothing
Response.Write msg
%>