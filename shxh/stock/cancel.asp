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
rst.Open "select * from 持股 where 持股人='"&username&"' and 数量>0 and 名称='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=066"
sopt=rst("出售")
snum=rst("数量")
sprice=rst("标价")
rst.Close
set rst=nothing
conn.BeginTrans
conn.Execute "update 持股 set 数量=0 where 持股人='"&username&"' and 名称='"&stock&"'"
if sopt=False then 	
	conn.Execute "update 用户 set 存款=存款+"&sprice*snum*0.98-200&" where 姓名='"&username&"'"
else
	conn.Execute "update 用户 set 存款=存款-200 where 姓名='"&username&"'"
end if	
conn.CommitTrans
conn.Close 
set conn=nothing
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>' text='FF0000'>
<table align=center width=100>
<tr align=center bgcolor=00ff00><td colspan=2>撤销收购请求</td></tr>
<tr align=center><td colspan=2><%=stock%></td></tr>
<tr><td>数量</td><td align=right><%=snum%></td></tr>
<tr><td>单价</td><td align=right><%=sprice%></td></tr>
<tr><td colspan=2><hr></td></tr>
<tr><td colspan=2 align=right><%=snum*sprice%></td></tr>
<%if sopt=false then%>
<tr><td>资费</td><td align=right><%=snum*sprice*0.02\1+200%></td></tr>
<tr><td>存款加</td><td align=right><%=snum*sprice*0.98\1-200%></td></tr>
<%else%>
<tr><td>资费</td><td align=right>200</td></tr>
<tr><td>存款减</td><td align=right>200</td></tr>
<%end if%>
<tr><td colspan=2 align=center><input type=button value='返回' onclick='javascript:history.back();'></tr>
</table>
</body>