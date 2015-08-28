<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
grade=request("grade")
if not isnumeric(grade) then Response.Redirect "../error.asp?id=024"
grade=clng(grade)
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>门派管理</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>姓名</td><td>门派</td><td>身份</td><td>等级</td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 等级="&grade,conn,1,3
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
rst.PageSize=20
if activepage>rst.PageCount then activepage=rst.PageCount
rst.AbsolutePage=activepage
for i=1 to rst.Pagesize
if rst.EOF or rst.BOF then exit for
msg=msg&"<tr><td><a href='showuser.asp?search="&rst("姓名")&"'>"&rst("姓名")&"</a></td><td>"&rst("门派")&"</td><td>"&rst("身份")&"</td><td>"&rst("等级")&"</td></tr>"
rst.MoveNext
next
msg=msg&"</table><form action='showbygr.asp?grade="&grade&"' method=post><table border=3 width='100%'><tr align=center bgcolor=cccccc>"
if activepage>1 then
msg=msg&"<td><a href='showbygr.asp?grade="&grade&"&activepage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='showbygr.asp?grade="&grade&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if activepage<rst.PageCount then
msg=msg&"<td><a href='showbygr.asp?grade="&grade&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='showbygr.asp?grade="&grade&"&activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=activepage size=4 value='"&activepage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form></body></html>"
rst.close
set rst=nothing
Response.Write msg
%>
