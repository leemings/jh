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
conn.open Application("sjjh_usermdb")
conn.Execute "update 用户 set 状态='监禁',事件原因='监禁:\n您在我们这里办理会员："&request.querystring("mess")&"，\n因为你不付款，我们监禁你的账号！!' where 姓名='"&request.querystring("name")&"'"
conn.Execute "update q set d='监禁' where a='"&request.querystring("name")&"'"
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：用户["&request.querystring("name")&"监禁完成！]');history.go(-1);</script>"
response.end

%>
