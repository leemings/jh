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
conn.Execute "delete * from q  where a='"&request.querystring("name")&"'"
conn.Execute "update 用户 set 状态='正常',事件原因='恭喜:\n您现在已经是我们的正式会员了，感谢您对我们的支持！' where 姓名='"&request.querystring("name")&"'"
request.querystring("name")
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：用户["&request.querystring("name")&"结算完成！]');history.go(-1);</script>"
response.end
%>
