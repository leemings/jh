<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then 
	Response.Write "<script language=javascript>top.location.replace('../error.asp?id=016');</script>"
	Response.End
end if	
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel=stylesheet href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' ><table border=1 width=100% align=center><tr bgcolor=FFFF00><td align=center colspan=2>药品</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select 名称,数量 from 物品 where 所有者='"&username&"' and 属性='药品' and 数量>0 order by 体力",conn
do while not(rst.EOF or rst.BOF)
	msg=msg&"<tr><td bgcolor='FFFF00'><a href='usecur.asp?cur="&rst("名称")&"' target=actfrm onmouseover="&chr(34)&"window.status='使用"&rst("名称")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("名称")&"</a></td><td align=right>"&rst("数量")&"</td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>