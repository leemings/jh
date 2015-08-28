<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
Response.Write "<script language=javascript>top.location.replace('../error.asp?id=016');</script>"
Response.End
end if
msg="<head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 background='../chatroom/bg1.gif'><table border=1 width=100% align=center><tr bgcolor=FFFF00><td align=center colspan=2>招式</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 招式,等级 from 攻击 where 姓名='"&username&"' order by 等级 desc",conn
do while not(rst.EOF or rst.BOF)
msg=msg&"<tr><td bgcolor='FFFF00'><a href='attnow.asp?attack="&rst("招式")&"' target=actfrm onmouseover="&chr(34)&"window.status='使用"&rst("招式")&"攻击';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("招式")&"</a></td><td align=right>"&rst("等级")&"</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>
