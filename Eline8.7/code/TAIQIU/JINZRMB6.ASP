<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员等级 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
hydj=rs("会员等级")
rs.close
set rs=nothing
conn.close
set conn=nothing
if hydj<2 then
	Response.Write "<script language=JavaScript>{alert('提示：会员等级为2级以上才可以银币换金币！');}</script>"
	Response.End
end if
%>
<title>银币转换金币</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="bg.gif">
<%my=sjjh_name
input=request.form("input1")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'yule6.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
if input/10<>int(input/10) then
 	Response.Write "<script language=JavaScript>{alert('请输入10的整数倍的数值！');location.href = 'yule6.asp';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 银币 from 用户 where 姓名='" & sjjh_name & "'",conn,2,2
if rs("银币")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('提示：你没有这么多的银币无法转换！');location.href = 'yule6.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 银币=银币-"&input&",金币=金币+" & int(input/10) & " where 姓名='"&sjjh_name&"'"
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"&sjjh_name&"把"&input&"个银币转换成金币:"& int(input/10)&"个！');location.href = 'yule6.asp';</script>"
%>
