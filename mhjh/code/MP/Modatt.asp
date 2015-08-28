<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
selcorp=Request.QueryString("selcorp")
if selcorp="" then selcorp="无"
username=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
msg="<head><link rel='stylesheet' href='../style.css'><title>各门武功</title></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><form name=form1><select name=sele1 onchange="&chr(34)&"javascript:location.replace('modatt.asp?selcorp='+document.form1.sele1.value);"&chr(34)&"><option value='无'>无门无派</option>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "门派",conn
do while not rst.EOF
corp=rst("门派")
if corp=selcorp then
msg=msg&"<option value='"&corp&"' selected>"&corp&"</option>"
else
msg=msg&"<option value='"&corp&"'>"&corp&"</option>"
end if
rst.MoveNext
loop
msg=msg&"</select></form></p><hr><table border=1 cellspacing=0 cellpadding=2  bordercolorlight=#993300 bordercolordark=#FFFFFF  align=center><tr align=center bgcolor=ffffff><td>招式</td><td>攻击</td><td>说明</td>"&opt&"</tr>"
rst.Close
rst.Open "select * from 招式 where 门派='"&selcorp&"' order by 基本攻击",conn
do while not (rst.EOF or rst.BOF)
attid=rst("id")
msg=msg&"<tr><td>"&rst("招式")&"</td><td align=right>"&rst("基本攻击")&"</td><td>"&rst("说明")&"</td>"&opt&"</tr>"
rst.MoveNext
loop
msg=msg&"</table></body>"
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>
