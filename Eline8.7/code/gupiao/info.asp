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
if stock="" then stock="『快乐江湖』"
msg="<head><link rel='stylesheet' href='../chat/readonly/style.css'><script language=javascript>setTimeout('location.reload();',300000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966' text='FF0000'><center><font color=red><h3>实时行情</h3></font><center><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center bgcolor=FFFF00><td>报价</td><td>数量</td></tr><tr align=center bgcolor=00FF00><td colspan=2><a href='sale.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='出售你手中的"&stock&"股份';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">出售</a></td></tr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select top 10 * from 持股 where 名称='"&stock&"' and 出售=True and 数量>0 order by 标价 desc,时间",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr align=right><td>"&rs("标价")&"</td><td>"&rs("数量")&"</td></tr>"
	rs.MoveNext
loop
rs.Close
msg=msg&"<tr align=center bgcolor=00FF00><td colspan=2><a href='buy.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='收购"&stock&"股份';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">收购</a></td></tr>"
rs.Open "select top 10 * from 持股 where 名称='"&stock&"' and 出售=False and 数量>0 order by 标价,时间",conn
do while not (rs.EOF or rs.BOF)
	msg=msg&"<tr align=right><td>"&rs("标价")&"</td><td>"&rs("数量")&"</td></tr>"
	rs.MoveNext
loop
rs.Close
rs.Open "select * from 持股 where 名称='"&stock&"' and 持股人='"&sjjh_name&"' and 数量>0",conn
if not rs.EOF then
	if rs("出售")=True then
		opt="出售"
	else
		opt="收购"
	end if 
	msg=msg&"<tr bgcolor=00ff00><td>操作</td><td>"&opt&"</td></tr><tr><td>"&rs("标价")&"</td><td>"&rs("数量")&"</td><tr><td colspan=2 align=center><a href='cancel.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='撤销你对"&stock&"的此次"&opt&"要求';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">撤销操作</a></td></tr></tr>"
end if
set rs=nothing
conn.Close
set conn=nothing
Response.Write msg&"</table></body>"
%>