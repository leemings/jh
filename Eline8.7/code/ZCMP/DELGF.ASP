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
id=trim(request("id"))
%><!--#include file="data.asp"--><%
sql="delete * from s where 申请人="&id
set rs=conn1.execute(sql)
conn1.close
set conn1=nothing
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>你成功把</title>
</head>

<body background="../JHIMG/BK_HC3W01.GIF">

<p align="center"><b><font size="5">删除成功</font></b></p>

<p align="center"><b><font size="5"><a href="gfsq.asp">返回</a></font></b></p>

</body>

</html>
