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
<title>金卡转换轻功♀wWw.happyjh.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<%my=sjjh_name
input=request.form("input")
if not isnumeric(input) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&input&"]输入错误，请使用数字！');location.href = 'qgzh.asp';}</script>"
	Response.End 
end if
input=int(abs(input))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员金卡,会员等级 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hy=rs("会员等级")
if rs("会员金卡")<input then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有这么多的金卡无法转换！');location.href = 'qgzh.asp';}</script>"
	response.end
end if
conn.execute "update 用户 set 会员金卡=会员金卡-" & input & ",轻功=轻功+"& int(input*50000)&" where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"&sjjh_name&"您转换法力:"&input*50000&"点！');location.href = 'qgzh.asp';</script>"
%>
