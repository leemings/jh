<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
corp=Request.QueryString("corp")
attid=Request.QueryString("attid")
opt=Request.QueryString("opt")
if not isnumeric(attid) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 招式 where id="&attid&" and 门派='"&corp&"'",conn
if opt="新增" then
	energy=10
	magic=2
	attack=10
else
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
	attackname=rst("招式")
	condition=rst("修习条件")
	energy=rst("消耗精力")
	magic=rst("消耗内力")
	attack=rst("基本攻击")
	special=rst("特效")
	introduce=rst("说明")
	attintro=rst("攻击说明")
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>武功设置</p><hr><form action=modatt3.asp method=post><input type=hidden name=attid value='"&attid&"'><table border=3 width=80% align=center><tr><td>门派</td><td><input type=text name=corp size=14 maxlength=7 value='"&corp&"'></td></tr><tr><td>招式名称</td><td><input type=text value='"&attackname&"'size=14 maxlength=7 name=attackname></td></tr><tr><td>修习条件</td><td><input type=text name=condition value="&chr(34)&condition&chr(34)&" size=50 maxlength=100></td></tr><tr><td>消耗精力</td><td><input type=text name=energy value="&chr(34)&energy&chr(34)&" size=5 maxlength=5></td></tr><tr><td>消耗内力</td><td><input type=text name=magic value="&chr(34)&magic&chr(34)&" size=5 maxlength=5></td></tr><tr><td>基本攻击</td><td><input type=text name=attack value="&chr(34)&attack&chr(34)&" size=5 maxlength=5></td></tr><tr><td>特效</td><td><input type=text size=4 maxlength=4 value='"&special&"'></td></tr><tr><td>说明</td><td><input type=text name=introduce value="&chr(34)&introduce&chr(34)&" size=50 maxlength=100></td></tr><tr><td>攻击说明</td><td><input type=text name=attintro value="&chr(34)&attintro&chr(34)&" size=30 maxlength=100></td></tr><tr align=center><td colspan=2><input type=submit name=opt value='"&opt&"'> <input type=button onclick=javascript:history.back(); value='返回'></td></tr></table></form></body>"
Response.Write msg
%>