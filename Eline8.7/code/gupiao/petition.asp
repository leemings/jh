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
stock=Request.QueryString("stock")
if stock="" then stock="��E�߽�����"
msg="<head><link rel='stylesheet' href='style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><h3>�깺ԭʼ��</h3><hr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.open Application("sjjh_usermdb")
set conn1=server.CreateObject("adodb.connection")
set rs1=server.CreateObject("adodb.recordset")
conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select ���� from �û� where ����='"&sjjh_name&"'",conn,2,2
umoney=rs("����")
rs.close
rs1.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn1
if rs1.EOF or rs1.BOF then 
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ݿ��ѯ���󣿣�');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
unum=umoney\rs1("���м�")
uallownum=unum
if uallownum>rs1("ʣ��ɷ�") then uallownum=rs1("ʣ��ɷ�")
	msg=msg&stock&"��ԭʼ�ɷݻ�ʣ��"&rs1("ʣ��ɷ�")&"��,��ӵ�н�������"&umoney&",<font color=FF0000>���깺"&uallownum&"��</font><form action='petitionnow.asp' method=post><table><tr><td>��Ʊ</td><td><input type=text value='"&stock&"' size=14 readonly name='stock'></td></tr><tr><td>����</td><td><input type=text value='"&uallownum&"' size=9 name='num'></td></tr><tr><td align=center colspan=2><input type=submit value=' �� �� '></td></tr></table></form>"
	rs1.Close
	conn.Close
	conn1.Close
	set rs=nothing
	set rs1=nothing
	set conn=nothing
	set conn1=nothing
	Response.Write msg
%>