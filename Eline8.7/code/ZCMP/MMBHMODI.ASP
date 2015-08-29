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
newwt=lcase(trim(request.form("wt")))
newda=lcase(trim(request.form("da")))
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
	Response.Write "<script Language=Javascript>alert('没有找到你要修改的密码保护资料！');window.close();</script>"
	response.end
end if
wt=rs("问题")
da=rs("答案")
rs.close
sql="update 密码保护 set 问题='" & newwt & "',答案='" &newda& "' where 用户="&id
set rs=conn1.execute(sql)
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>密码保护资料修改成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<p align="center"><b><font size="5">修改成功</font></b></p>
<p align="center">用户ID：<%=id%></p>
<p align="center">提示问题：<%=wt%>-------<%=newwt%></p>
<p align="center">答案：<%=da%>-------<%=newda%></p>
<p align="center"><a href="mmbhcx.asp">返回</a></p>
<p align="center"><font color="#FF0000" size="2"><b>『E线江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:一线天</b></font></p>
</body>
</html>