<%
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
co=Request("corp")
if instr(co,"'")<>0 then Response.Redirect "../error.asp?id=024"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>门派管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>姓名</td><td>门派</td><td>身份</td><td>等级</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 门派='"&co&"'",conn,1,3
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
rst.PageSize=20
if activepage>rst.PageCount then activepage=rst.PageCount
rst.AbsolutePage=activepage
for i=1 to rst.Pagesize
	if rst.EOF or rst.BOF then exit for
	msg=msg&"<tr><td><a href='showuser.asp?search="&rst("姓名")&"'>"&rst("姓名")&"</a></td><td>"&rst("门派")&"</td><td>"&rst("身份")&"</td><td>"&rst("等级")&"</td></tr>"
	rst.MoveNext
next
msg=msg&"</table><form action='showbyco.asp?corp="&co&"' method=post><table border=3 width='100%'><tr align=center bgcolor=cccccc>"
if activepage>1 then
	msg=msg&"<td><a href='showbyco.asp?corp="&co&"&activepage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='showbyco.asp?corp="&co&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
	msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if activepage<rst.PageCount then
	msg=msg&"<td><a href='showbyco.asp?corp="&co&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='showbyco.asp?corp="&co&"&activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
	msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=activepage size=4 value='"&activepage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form></body></html>"
Response.Write msg
%>