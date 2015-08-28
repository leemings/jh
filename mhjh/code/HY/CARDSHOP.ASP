<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<p align="center">
<font size="4" face="宋体">
&nbsp;道具卡片</font>
<hr>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from cardshop order by cprice",conn
do while not rst.EOF
msg=msg&"<tr><td><a href='buycard.asp?cid="&rst("id")&"' onmouseover="&chr(34)&"window.status='购买"&rst("cname")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("cname")&"</a></td><td>"&rst("cespecial")&"</td><td align=right>"&rst("cprice")&"</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if yx8_mhjh_fellow=true then assert="（只对会员开放）"
Response.Write "<html><head><link rel=stylesheet href='../style.css'><title>道具卡片</title></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><table width='80%' align=center border='1' cellspacing='0' cellpadding='1' bordercolorlight=#000000 bordercolordark=#FFFFFF><tr align=center bgcolor=ffffff><td>道具名称</td><td>作用</td><td>价格</td></tr>"&msg&"</table></body></html>"
%>
<p align="center"><b><%=assert%></b></p>
