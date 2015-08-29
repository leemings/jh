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
%>
<title>法力转换金卡♀wWw.51eline.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<%my=sjjh_name
input=request.form("input1")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'flzh.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
if input/100000<>int(input/100000) then
 	Response.Write "<script language=JavaScript>{alert('请输入100000的整数倍的数值！');location.href = 'flzh.asp';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力 from 用户 where 姓名='" & sjjh_name & "'",conn,2,2
if rs("法力")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('提示：你没有这么多的法力无法转换！');location.href = 'flzh.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 法力=法力-"&input&",会员金卡=会员金卡+" & int(input/100000) & " where 姓名='"&sjjh_name&"'"
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"&sjjh_name&"把"&input&"点法力转换成金卡:"& int(input/100000)&"元！');location.href = 'flzh.asp';</script>"
%>
