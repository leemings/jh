<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
stock=Request.form("stock")
unum=Request.Form("unum")
uprice=Request.Form("uprice")
if instr(stock,"'")<>0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if not(isnumeric(unum) and isnumeric(uprice)) then 
	Response.Write "<script Language=Javascript>alert('��ʾ����������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
unum=clng(unum)
if unum<1 then unum=1
uprice=clng(uprice)
if uprice<1 then uprice=1
uallnum=unum
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
set conn1=server.CreateObject("adodb.connection")
conn1.open Application("sjjh_usermdb")
rs.Open "select * from �ֹ� where �ֹ���='"&sjjh_name&"' and ����='"&stock&"' and ��Ȩ>="&unum,conn
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ���,��û���㹻�����Ĵ��ֹ�Ʊ�ɹ����ۣ�');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.Close
rs.Open "select * from �ֹ� where ����=False and ����>0 and ����='"&stock&"' and ���>="&uprice&" order by ʱ��"
do while not rs.EOF
	num=rs("����")
	st=rs("�ֹ���")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
		unum=0 
	end if 
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&num&",����=����-"&num&" where �ֹ���='"&st&"' and ����='"&stock&"'"
	if unum=0 then exit do
	rs.MoveNext
loop
rs.Close
set rs=nothing
uright=uallnum-unum
if uright<>0 then conn.Execute "update ��Ʊ set ����ɽ���=�ּ�,�ּ�="&uprice&" where ��Ʊ����='"&stock&"'"
conn.Execute "update �ֹ� set ��Ȩ=��Ȩ-"&uright&",����=True,���="&uprice&",����="&unum&" where �ֹ���='"&sjjh_name&"' and ����='"&stock&"'"
conn1.Execute "update �û� set ����=����+"&uright*uprice&" where ����='"&sjjh_name&"'"
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../chat/readonly/style.css'><script language=javascript>top.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966'><h2>ʵʱ����</h2><hr><p align=center>������������Ӻ��Զ�����<br><a href='javascript:history.go(-2)'> �� �� </a></p></body>"
%>