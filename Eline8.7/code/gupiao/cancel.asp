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
if stock="" or instr(stock,"'") then
	Response.Write "<script Language=Javascript>alert('��ʾ����������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
set conn1=server.CreateObject("adodb.connection")
conn1.open Application("sjjh_usermdb")
rs.Open "select * from �ֹ� where �ֹ���='"&sjjh_name&"' and ����>0 and ����='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ�������ǰ��û�жԴ˹�Ʊ���в���Ҫ�����𣿣�');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
sopt=rs("����")
snum=rs("����")
sprice=rs("���")
rs.Close
set rs=nothing
conn.BeginTrans
conn.Execute "update �ֹ� set ����=0 where �ֹ���='"&sjjh_name&"' and ����='"&stock&"'"
if sopt=False then 
	conn1.Execute "update �û� set ����=����+"&sprice*snum*0.98-200&" where ����='"&sjjh_name&"'"
else
	conn1.Execute "update �û� set ����=����-200 where ����='"&sjjh_name&"'"
end if 
conn.CommitTrans
conn.Close
set conn=nothing
conn1.close
set conn1=nothing
%>
<head><link rel='stylesheet' href='../chat/readonly/style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'>
<table align=center width=100>
<tr align=center bgcolor=00ff00><td colspan=2>�����չ�����</td></tr>
<tr align=center><td colspan=2><%=stock%></td></tr>
<tr><td>����</td><td align=right><%=snum%></td></tr>
<tr><td>����</td><td align=right><%=sprice%></td></tr>
<tr><td colspan=2><hr></td></tr>
<tr><td colspan=2 align=right><%=snum*sprice%></td></tr>
<%if sopt=false then%>
<tr><td>�ʷ�</td><td align=right><%=snum*sprice*0.02\1+200%></td></tr>
<tr><td>������</td><td align=right><%=snum*sprice*0.98\1-200%></td></tr>
<%else%>
<tr><td>�ʷ�</td><td align=right>200</td></tr>
<tr><td>������</td><td align=right>200</td></tr>
<%end if%>
<tr><td colspan=2 align=center><input type=button value='����' onclick='javascript:history.back();'></tr>
</table>
</body>