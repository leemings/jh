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
msg="<tr bgcolor=00FF00><td>名称</td><td>说明</td><td>版主</td><td width=40><a href='upboard2.asp?opt=新增&id=-1'>新增</a></td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "版块",conn
do while not rst.EOF
	msg=msg&"<tr><td>"&rst("名称")&"</td><td>"&rst("说明")&"</td><td>"&rst("版主")&"</td><td><a href='upboard2.asp?opt=修改&id="&rst("id")&"'>修改</a><br><a href='upboard2.asp?opt=删除&id="&rst("id")&"'>删除</a></td></tr>"
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><p align=center><font color=ff0000 size=5>论坛管理</font></p><table width='80%' align=center border=3>"&msg&"</table></body>"
%>