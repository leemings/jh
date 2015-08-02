<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
selcorp=Request.QueryString("selcorp")
if selcorp="" then selcorp="无"
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'>.</head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><form name=form1><select name=sele1 onchange="&chr(34)&"javascript:location.replace('modatt1.asp?selcorp='+document.form1.sele1.value);"&chr(34)&"><option value='无'>无门无派</option>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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
if usercorp="官府" and usergrade=Application("Ba_jxqy_allright") then opt="<td><a href='modatt2.asp?opt=新增&corp="&selcorp&"&attid=-1'>新&nbsp;增&nbsp;招&nbsp;式</a></td>"
msg=msg&"</select></form></p><hr><table border=3 width=80% align=center><tr align=center bgcolor=FFFF00><td>招式</td><td>攻击</td><td>说明</td>"&opt&"</tr>"
rst.Close
rst.Open "select * from 招式 where 门派='"&selcorp&"' order by 基本攻击",conn
do while not (rst.EOF or rst.BOF)
	attid=rst("id")
	if usercorp="官府" and usergrade=Application("Ba_jxqy_allright") then opt="<td><a href='modatt2.asp?opt=修改&corp="&selcorp&"&attid="&attid&"'>修改</a> | <a href='modatt2.asp?opt=删除&corp="&selcorp&"&attid="&attid&"'>删除</a></td>"
	msg=msg&"<tr><td>"&rst("招式")&"</td><td align=right>"&rst("基本攻击")&"</td><td>"&rst("说明")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
msg=msg&"</table></body>"
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>