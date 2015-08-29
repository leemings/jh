<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM q where d='未付'",conn,3,3
do while not rs.eof 
	conn.Execute "update 用户 set 事件原因='警告:\n您在我们这里办理会员："&rs("e")&"，\n因为你不付款，我们监禁你的账号！!' where 姓名='"&rs("a")&"'"
	rs.movenext
loop
rs.close
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：所有用户警告完成！]');history.go(-1);</script>"
response.end
%>
