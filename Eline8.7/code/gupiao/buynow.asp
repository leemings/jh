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
unum=int(abs(clng(unum)))
if unum<1 then unum=1
uprice=int(abs(clng(uprice)))
if uprice<1 then uprice=1
umoney=uprice*unum
uallnum=unum
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ���û�д��ֹ�Ʊ��');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if rs("��ͨ������")<unum then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ�����Ҫ�������Ѿ���������ͨ������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open Application("sjjh_usermdb")
rs1.open "select * from �û� where ����='"&sjjh_name&"' and ����>="&umoney,conn1,2,2
if rs1.EOF or rs1.BOF then 
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ������Ǯ���󲻹�Ӵ��');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs1.Close
rs.Open "select * from �ֹ� where �ֹ���='"&sjjh_name&"' and ����='"&stock&"' and ����>0",conn,2,2
if not (rs.EOF or rs.BOF) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ����޷�֧�ֶ�ͬһֻ��Ʊ�������β��������ȳ����ϴεĲ���Ҫ��');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
rs.Close
rs.Open "select * from �ֹ� where ����=True and ����>0 and ����='"&stock&"' and ���<="&uprice&" order by ʱ��",conn
do while not rs.EOF
	num=rs("����")
	st=rs("�ֹ���")
	if num<unum then
		unum=unum-num
	elseif num>=unum then
		num=unum
	unum=0 
	end if 
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ-"&num&",����=����-"&num&" where �ֹ���='"&st&"' and ����='"&stock&"'"
	conn1.Execute "update �û� set ����=����+"&num*uprice&" where ����='"&st&"'"
	if unum=0 then exit do
	rs.MoveNext
loop
rs.Close
uright=uallnum-unum
if uright<>0 then conn.Execute "update ��Ʊ set ����ɽ���=�ּ�,�ּ�="&uprice&" where ��Ʊ����='"&stock&"'"
rs.Open "select * from �ֹ� where �ֹ���='"&sjjh_name&"' and ����='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	conn.Execute "insert into �ֹ�(�ֹ���,ʱ��,����,��Ȩ,����,���,����) values('"&sjjh_name&"',now(),'"&stock&"',"&uright&",False,"&uprice&","&unum&")"
else
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&uright&",����=False,���="&uprice&",����="&unum&" where �ֹ���='"&sjjh_name&"' and ����='"&stock&"'"
end if
conn1.Execute "update �û� set ����=����-"&umoney&" where ����='"&sjjh_name&"'"
set rs=nothing
set rs1=nothing
conn.Close
conn1.Close
set conn=nothing
set conn1=nothing
Response.Write "<head><link rel='stylesheet' href='../chat/readonly/style.css'><script language=javascript>parent.pricefrm.location.reload();</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966'><h2>ʵʱ����</h2><hr><table width='100%'><tr><td colspan=2 align=center bgcolor=00ff00>"&stock&"</td></tr><tr><td >����</td><td align=right>"&uprice&"</td></tr><tr><td>Ҫ��</td><td align=right>"&uallnum&"</td></tr><tr><td>�ɽ�</td><td align=right>"&uright&"</td></tr><tr><td>���</td><td align=right>"&uright*uprice&"</td></tr><tr><td>Ѻ��</td><td align=right>"&unum*uprice&"</td><tr><td>ʵ��</td><td align=right>"&umoney&"</td></tr></tr></table><p align=center><a href='javascript:history.go(-2)'> �� �� </a></p></body>"
%>