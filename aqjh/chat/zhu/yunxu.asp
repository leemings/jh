<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->

<%Response.charset="gb2312"%><%
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 组队,操作时间,组名 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("组名")<>"" and rs("组名")<>"无" then 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你正在组内，请离开组后再进行此操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if

if rs("组队")=true then
	conn.Execute "update 用户 set 组队=false,操作时间=now() where 姓名='"&aqjh_name&"'"	
		diaox="禁止了组队，希望捣乱的组长不要骚扰他，乱加组员会扣经验的"

else
	diaox="打开组队状态，现在可以叫组长加你了"
	conn.Execute "update 用户 set 组队=true,操作时间=now() where 姓名='"&aqjh_name&"'"	
end if
	Response.Write "<script language=JavaScript>{alert('组队状态切换成功！');location.href = 'javascript:history.go(-1)';}</script>"

        rs.close

conn.close
set rs=nothing
set conn=nothing
Application.Lock

diaox="<font color=#ff0000>【组队消息】</font> <font color=#ff00ff>" & aqjh_name & "</font>"&diaox
call showchat(diaox)

%>

</body>
</html>