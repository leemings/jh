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
uallnum=unum
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"' and ��Ȩ>="&unum,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=052"
rst.Close
conn.BeginTrans 
rst.Open "select * from �ֹ� where ����=False and ����>0 and ����='"&stock&"' and ���>="&uprice&" order by ʱ��"
do while not rst.EOF
	num=rst("����")
	st=rst("�ֹ���")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
		unum=0	
	end if	
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&num&",����=����-"&num&" where �ֹ���='"&st&"' and ����='"&stock&"'"
	if unum=0 then exit do
	rst.MoveNext 
loop
rst.Close
set rst=nothing
uright=uallnum-unum
if uright<>0 then conn.Execute "update ��Ʊ set ����ɽ���=�ּ�,�ּ�="&uprice&" where ��Ʊ����='"&stock&"'"
conn.Execute "update �ֹ� set ��Ȩ=��Ȩ-"&uright&",����=True,���="&uprice&",����="&unum&" where �ֹ���='"&username&"' and ����='"&stock&"'"
conn.Execute "update �û� set ���=���+"&uright*uprice&" where ����='"&username&"'"
conn.CommitTrans
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>top.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><h2>ʵʱ����</h2><hr><p align=center>������������Ӻ��Զ�����<br><a href='javascript:history.go(-2)'> �� �� </a></p></body>"
%>