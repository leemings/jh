<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" or instr(stock,"'") then Response.Redirect "../error.asp?id=100"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����>0 and ����='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=066"
sopt=rst("����")
snum=rst("����")
sprice=rst("���")
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "update �ֹ� set ����=0 where �ֹ���='"&username&"' and ����='"&stock&"'"
if sopt=False then 	
	conn.Execute "update �û� set ���=���+"&sprice*snum*0.98-200&" where ����='"&username&"'"
else
	conn.Execute "update �û� set ���=���-200 where ����='"&username&"'"
end if	
conn.CommitTrans
conn.Close 
set conn=nothing
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>' text='FF0000'>
<table align=center width=100>
<tr align=center bgcolor=00ff00><td colspan=2>�����չ�����</td></tr>
<tr align=center><td colspan=2><%=stock%></td></tr>
<tr><td>����</td><td align=right><%=snum%></td></tr>
<tr><td>����</td><td align=right><%=sprice%></td></tr>
<tr><td colspan=2><hr></td></tr>
<tr><td colspan=2 align=right><%=snum*sprice%></td></tr>
<%if sopt=false then%>
<tr><td>�ʷ�</td><td align=right><%=snum*sprice*0.02\1+200%></td></tr>
<tr><td>����</td><td align=right><%=snum*sprice*0.98\1-200%></td></tr>
<%else%>
<tr><td>�ʷ�</td><td align=right>200</td></tr>
<tr><td>����</td><td align=right>200</td></tr>
<%end if%>
<tr><td colspan=2 align=center><input type=button value='����' onclick='javascript:history.back();'></tr>
</table>
</body>