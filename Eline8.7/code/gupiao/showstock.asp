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
msg="<head><link rel='stylesheet' href='style.css'><script language=javascript>function petition(stock){parent.infofrm.location.replace('petition.asp?stock='+stock);}</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><center><font color=red><h3>"&stock&"</h3></font></center><table width=80% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	msg=msg&"<tr><td>��Ҫ�ҵĹ�Ʊû���ҵ�</td></tr>"
else
	for i=1 to rs.Fields.Count -1
		msg=msg&"<tr><td bgcolor=00ff00>"&rs.Fields(i).name&"</td><td align=right>"&rs.Fields(i).value&"</td></tr>"
	next
	if rs("ʣ��ɷ�")<>0 then msg=msg&"<tr bgcolor=ffff00><td colspan=2 align=center><input type=button onclick="&chr(34)&"javascript:petition('"&stock&"');"&chr(34)&" value=' �� �� '></td></tr>"
end if
msg=msg&"</table><hr bgcolor=red size=1><table width=80% align=center border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF ><tr><td align=center colspan=2>��Ҫ�ֹ���</td></tr>"
rs.Close
rs.Open "select top 10 * from �ֹ� where ����='"&stock&"' and ��Ȩ>0 order by ��Ȩ desc",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr><td>"&rs("�ֹ���")&"</td><td align=right>"&rs("��Ȩ")&"��</td></tr>"
	rs.MoveNext
loop
msg=msg&"</table></body>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write msg
%>