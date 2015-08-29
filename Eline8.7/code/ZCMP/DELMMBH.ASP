<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then
	Response.Write "<script Language=Javascript>alert('你不是站长？！');window.close();</script>"
	response.end
end if
id=lcase(trim(request("id")))
%>
<!--#include file="data.asp"-->
<%
sql="select * from 密码保护 where 用户="&id
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('没有找到你要删除的密码保护资料！');window.close();</script>"
	response.end
end if
rs.close
sql="delete * from 密码保护 where 用户="&id
set rs=conn1.execute(sql)
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>删除密码保护</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="../hc3w_img/bg.gif">
<p align="center">删除<%=id%>的密码保护资料成功 </p>
<p align="center"><a href="mmbhcx.asp">返回</a> 
</p>
</body>
</html>

