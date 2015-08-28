<%
if session("yx8_mhjh_username")="" then
response.write("SORRY！您没有<a href=../index.asp>登录</a>系统。")
else
%>
<%my=session("yx8_mhjh_username")%>
<!--#include file="data.asp"-->
<%
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
pass=rs("帐号")
roomname=my
username=my
password=pass

if username="" or password="" or roomname="" then
	response.write "请输入完整信息！"
	response.end
end if
dim conntmp,DB_name
DB_name="DBQ="+server.mappath("../21ex/fangzi.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conntmp=server.createobject("ADODB.CONNECTION")
conntmp.open DB_name
strSQL = "SELECT * FROM room WHERE roomname='" & roomname & "'"
Set user=Server.CreateObject("ADODB.Recordset")
user.open strSQL,conntmp,3,2
If user.eof and user.bof then
	user.addnew
	user("username")=username
	user("roomname")=roomname
	user("password")=password
	user("state") = False
	user("sj") = Now()
	user("de") =false
	user.update
	response.write "已经成功的建立房间请<br>"
	'response.Write "<a href="enter.asp?roomname=" &roomname&"&sf=sf&name=" & username & "">进入聊天室</a>"
	session("sysop")=True
	response.write "<a href=""xiaowu.asp"">返回</a>"
	else
	response.write "已经存在相同的房间了！"
	response.write "<a href=""xiaowu.asp"">返回</a>"
	response.end
End If
	user.close
	conntmp.close
	set user = nothing
	set conntmp = nothing
	response.end
End If
%>