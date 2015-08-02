<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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
if Application("Ba_jxqy_fellow")=true then assert="（只对会员开放）"
Response.Write "<html><head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false bgcolor='"&bgcolor&"' background='"&bgimage&"'><p align=center><font color=0000ff size=5>道具店"&assert&"</font></p><hr><table width='80%' align=center border=3><tr align=center bgcolor=ffff00><td>道具名称</td><td>作用</td><td>价格</td></tr>"&msg&"</table><p align=center><input type=button value=' 返 回 ' onClick=javascript:location.href='../street/street.asp'; id=button1 name=button1> <input type=button value=' 关 闭 ' onClick=javascript:window.close(); id=button1 name=button1></p></body></html>"
%>