<%
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=50 and usercorp="官府") then Response.Redirect "../error.asp?id=046"
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"

set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 答题  where id="&id,conn
if opt="新增" then
	question=""
	answer=""
	ti=""
	money=""
	qiang=false
else
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
	question=rst("问题")
	answer=rst("答案")
	ti=rst("提供者")
	money=rst("奖金")
	qiang=rst("抢答")
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if qiang=true then
	qiang="<select name=qiang><option value='true' selected>抢答题</option><option value='false' >正常题</option></select>"
else
	qiang="<select name=qiang><option value='true' >抢答题</option><option value='false'  selected>正常题</option></select>"
end if
msg=msg&"<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>答题管理</p><hr><form action=updateti.asp method=post><input type=hidden name='schid' value='"&id&"'><table border=3 width=90% align=center><tr><td>问题</td><td><input type=text name=question size=50 maxlength=255 value='"&question&"'></td></tr><tr><td>答案</td><td><input type=text value='"&answer&"'size=20 maxlength=30 name=answer></td></tr><tr><td>提供者</td><td><input type=text name=ti value="&chr(34)&ti&chr(34)&" size=10 maxlength=20></td></tr><tr><td>奖金</td><td><input type=text name=money value="&chr(34)&money&chr(34)&" size=10 maxlength=10></td></tr><tr><td>抢答</td><td>"&qiang&"</td></tr><tr align=center><td colspan=2><input type=submit name=submit value='"&opt&"'> <input type=button onclick=javascript:history.back(); value='返回'></td></tr></table></body>"
Response.Write msg
%>