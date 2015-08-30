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
<title>金币转换银币♀wWw.happyjh.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="bg.gif">
<%my=sjjh_name
input=request.form("input")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'yule6.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 金币,会员等级 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hy=rs("会员等级")
if rs("金币")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有这么多的金币无法转换！');location.href = 'yule6.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 金币=金币-" & input & ",银币=银币+"& int(input*10)&" where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"&sjjh_name&"您转换银币:"&input*10&"个！');location.href = 'yule6.asp';</script>"
%>
