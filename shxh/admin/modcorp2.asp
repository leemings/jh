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
corpid=Request.QueryString("corpid")
opt=Request.QueryString("opt")
if not isnumeric(corpid) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 门派  where id="&corpid,conn
if opt="新增" then
	corp=""
	earnings=0
	introduce=""
	condition="True"
	chaton=false
else
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
	corp=rst("门派")
	earnings=rst("工资系数")
	introduce=rst("简介")
	condition=rst("加入条件")
	chaton=rst("chaton")
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if chaton=true then
	chaton="<select name=chaton><option value='true' selected>开设聊天室</option><option value='false' >关闭聊天室</option></select>"
else
	chaton="<select name=chaton><option value='true' >开设聊天室</option><option value='false'  selected>关闭聊天室</option></select>"
end if
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>门派管理</p><hr><form action=modcorp3.asp method=post><input type=hidden name=corpid value='"&corpid&"'><table border=3 width=80% align=center><tr><td>门派</td><td><input type=text name=corp size=14 maxlength=7 value='"&corp&"'></td></tr><tr><td>工资系数</td><td><input type=text value='"&earnings&"'size=5 maxlength=9 name=earnings></td></tr><tr><td>简介</td><td><input type=text name=introduce value="&chr(34)&introduce&chr(34)&" size=50 maxlength=100></td></tr><tr><td>加入条件</td><td><input type=text name=condition value="&chr(34)&condition&chr(34)&" size=25 maxlength=50></td></tr><tr><td>聊天室</td><td>"&chaton&"</td></tr><tr align=center><td colspan=2><input type=submit name=opt value='"&opt&"'> <input type=button onclick=javascript:history.back(); value='返回'></td></tr></table></form></body>"
Response.Write msg
%>