<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" then stock="梦想成真"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=Session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('location.reload();',300000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><center><font color=red><h3>实时行情</h3></font><center><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center bgcolor=FFFF00><td>报价</td><td>数量</td></tr><tr align=center bgcolor=00FF00><td colspan=2><a href='sale.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='出售你手中的"&stock&"股份';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">出售</a></td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select top 10 * from 持股 where 名称='"&stock&"' and 出售=True and 数量>0  order by  标价 desc,时间",conn
do while not (rst.EOF or rst.BOF) 
	msg=msg&"<tr align=right><td>"&rst("标价")&"</td><td>"&rst("数量")&"</td></tr>"
	rst.MoveNext
loop
rst.Close
msg=msg&"<tr align=center bgcolor=00FF00><td colspan=2><a href='buy.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='收购"&stock&"股份';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">收购</a></td></tr>"
rst.Open "select top 10 * from 持股 where 名称='"&stock&"' and 出售=False and 数量>0  order by  标价,时间",conn
do while not (rst.EOF or rst.BOF) 
	msg=msg&"<tr align=right><td>"&rst("标价")&"</td><td>"&rst("数量")&"</td></tr>"
	rst.MoveNext 
loop
rst.Close
rst.Open "select * from 持股 where 名称='"&stock&"' and 持股人='"&username&"' and 数量>0",conn
if not rst.EOF then
	if rst("出售")=True then 
		opt="出售"
	else
		opt="收购"
	end if	
	msg=msg&"<tr bgcolor=00ff00><td>操作</td><td>"&opt&"</td></tr><tr><td>"&rst("标价")&"</td><td>"&rst("数量")&"</td><tr><td colspan=2 align=center><a href='cancel.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='撤销你对"&stock&"的此次"&opt&"要求';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">撤销操作</a></td></tr></tr>"
end if
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</table></body>"
%>