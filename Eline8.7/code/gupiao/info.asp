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
if stock="" then stock="�����ֽ�����"
msg="<head><link rel='stylesheet' href='../chat/readonly/style.css'><script language=javascript>setTimeout('location.reload();',300000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><center><font color=red><h3>ʵʱ����</h3></font><center><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center bgcolor=FFFF00><td>����</td><td>����</td></tr><tr align=center bgcolor=00FF00><td colspan=2><a href='sale.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='���������е�"&stock&"�ɷ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">����</a></td></tr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select top 10 * from �ֹ� where ����='"&stock&"' and ����=True and ����>0 order by ��� desc,ʱ��",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr align=right><td>"&rs("���")&"</td><td>"&rs("����")&"</td></tr>"
	rs.MoveNext
loop
rs.Close
msg=msg&"<tr align=center bgcolor=00FF00><td colspan=2><a href='buy.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='�չ�"&stock&"�ɷ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�չ�</a></td></tr>"
rs.Open "select top 10 * from �ֹ� where ����='"&stock&"' and ����=False and ����>0 order by ���,ʱ��",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr align=right><td>"&rs("���")&"</td><td>"&rs("����")&"</td></tr>"
	rs.MoveNext
loop
rs.Close
rs.Open "select * from �ֹ� where ����='"&stock&"' and �ֹ���='"&sjjh_name&"' and ����>0",conn
if not rs.EOF then
	if rs("����")=True then
		opt="����"
	else
		opt="�չ�"
	end if 
	msg=msg&"<tr bgcolor=00ff00><td>����</td><td>"&opt&"</td></tr><tr><td>"&rs("���")&"</td><td>"&rs("����")&"</td><tr><td colspan=2 align=center><a href='cancel.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='�������"&stock&"�Ĵ˴�"&opt&"Ҫ��';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��������</a></td></tr></tr>"
end if
set rs=nothing
conn.Close
set conn=nothing
Response.Write msg&"</table></body>"
%>