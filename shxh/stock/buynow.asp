<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
stock=Request.form("stock")
unum=Request.Form("unum")
uprice=Request.Form("uprice")
if instr(stock,"'")<>0 then Response.Redirect "../error.asp?id=024"
if not(isnumeric(unum) and isnumeric(uprice)) then Response.Redirect "../error.asp?id=024"
unum=clng(unum) 
if unum<1 then unum=1
uprice=clng(uprice)
if uprice<1 then uprice=1
umoney=uprice*unum
uallnum=unum
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.open "select * from �û� where ����='"&username&"' and ���>="&umoney,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"' and ����>0",conn
if not (rst.EOF or rst.BOF) then  Response.Redirect "../error.asp?id=067"
rst.Close
rst.Open "select * from �ֹ� where ����=True and ����>0 and ����='"&stock&"' and ���<="&uprice&" order by ʱ��"
conn.BeginTrans
do while not rst.EOF
	num=rst("����")
	st=rst("�ֹ���")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
		unum=0	
	end if	
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ-"&num&",����=����-"&num&" where �ֹ���='"&st&"' and ����='"&stock&"'"
	conn.Execute "update �û� set ���=���+"&num*uprice&" where ����='"&st&"'"
	if unum=0 then exit do
	rst.MoveNext 
loop
rst.Close
uright=uallnum-unum
if uright<>0 then conn.Execute "update ��Ʊ set ����ɽ���=�ּ�,�ּ�="&uprice&" where ��Ʊ����='"&stock&"'"
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"'",conn
if rst.EOF or rst.BOF then
	conn.Execute "insert into �ֹ�(�ֹ���,ʱ��,����,��Ȩ,����,���,����) values('"&username&"',now(),'"&stock&"',"&uright&",False,"&uprice&","&unum&")"
else
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&uright&",����=False,���="&uprice&",����="&unum&" where �ֹ���='"&username&"' and ����='"&stock&"'"
end if
conn.Execute "update �û� set ���=���-"&umoney&" where ����='"&username&"'"
conn.CommitTrans
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>parent.pricefrm.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><h2>ʵʱ����</h2><hr><table width='100%'><tr><td colspan=2 align=center bgcolor=00ff00>"&stock&"</td></tr><tr><td >����</td><td align=right>"&uprice&"</td></tr><tr><td>Ҫ��</td><td align=right>"&uallnum&"</td></tr><tr><td>�ɽ�</td><td align=right>"&uright&"</td></tr><tr><td>���</td><td align=right>"&uright*uprice&"</td></tr><tr><td>Ѻ��</td><td align=right>"&unum*uprice&"</td><tr><td>ʵ��</td><td align=right>"&umoney&"</td></tr></tr></table><p align=center><a href='javascript:history.go(-2)'> �� �� </a></p></body>"
%>