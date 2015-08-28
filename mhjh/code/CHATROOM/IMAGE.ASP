<%
msg="<html><head><link rel='stylesheet' href='css3.css'></head><body  background='bg1.gif' ><p align=center><font color=0000FF size=4>ÌùÍ¼</font><hr></p>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "img",conn
do while not rst.EOF
img=Trim(rst("filename"))
msg=msg&"<a href='#' onclick="&chr(34)&"javascript:parent.talkfrm.addtalk('[img]"&img&"[/img]');"&chr(34)&" onmouseover="&chr(34)&"window.status='"&img&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='"&img&"'><img src='../chatroom/image/image/"&img&"' border=0></a>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</body></html>"
%>
