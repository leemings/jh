<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" then stock="梦想成真工作室"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><h3>申购原始股</h3><hr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 存款 from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=016"
umoney=rst("存款")
rst.Close
rst.Open "select * from 股票 where 股票名称='"&stock&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
unum=umoney\rst("发行价")
uallownum=unum
if uallownum>rst("剩余股份") then uallownum=rst("剩余股份")
msg=msg&stock&"的原始股份还剩余"&rst("剩余股份")&"股,你拥有银行存款"&umoney&",<font color=FF0000>可申购"&uallownum&"股</font><form action='petitionnow.asp' method=post><table><tr><td>股票</td><td><input type=text value='"&stock&"' size=14 readonly name='stock'></td></tr><tr><td>数量</td><td><input type=text value='"&uallownum&"' size=9 name='num'></td></tr><tr><td align=center colspan=2><input type=submit value='  申 购 '></td></tr></table></form>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>