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
msg="<head><link rel=stylesheet href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' ><table border=1 width=100% align=center><tr bgcolor=FFFF00><td align=center colspan=2>ҩƷ</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select ����,���� from ��Ʒ where ������='"&username&"' and ����='ҩƷ' and ����>0 order by ����",conn
do while not(rst.EOF or rst.BOF)
	msg=msg&"<tr><td bgcolor='FFFF00'><a href='usecur.asp?cur="&rst("����")&"' target=actfrm onmouseover="&chr(34)&"window.status='ʹ��"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("����")&"</a></td><td align=right>"&rst("����")&"</td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>