<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
msg="<head><title>��������</title><link rel=stylesheet href='../../chatroom/css.css'></head><body  background=../../chatroom/bg1.gif oncontextmenu='self.event.returnValue=false;'><center><b>��������</b><br><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center ><tr align=center bgcolor=EEF7F7><td >����</td><td>������</td><td>������</td></tr>"
%><!--#include file="data.asp"--><%  
sqlstr="select lei,username,attack from myanimal where outtime=false order by attack desc"
rs.open sqlstr,connb,1,2
if not (rs.EOF or rs.BOF) then
rs.PageSize=12
if acpage>rs.PageCount then acpage=rs.PageCount
rs.AbsolutePage=acpage
for i=1 to rs.PageSize
if rs.EOF then exit for
		msg=msg&"<tr><td>"&rs("lei")&"</td><td>"&rs("username")&"</td><td>"&rs("attack")&"</td></tr>"
		rs.Movenext
next
msg=msg&"</table><form action='animaltop.asp' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=#EEF7F7><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='animaltop.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='animaltop.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
end if
if acpage<rs.PageCount then
msg=msg&"<td><a href='animaltop.asp?acpage="&acpage+1&"'  onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='animaltop.asp?acpage="&rs.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
end if
msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rs.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
else
msg=msg&"</table>"
end if
rs.Close
set rs=nothing
connb.Close
set connb=nothing
msg=msg&"<p align=center></p></body>"
Response.Write msg
%>