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
if sjjh_grade<10 or sjjh_name<>application("sjjh_user") then
	Response.Write "<b>[操作失败]</b><p>您没有权限！"
	Response.End
end if
downvalue=Trim(Request.Form("downvalue"))
if IsNumeric(downvalue)=false then
	Response.write "<b>[操作失败]</b><p>必须输入数值！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.end
end if
if downvalue="" then downvalue=0
if downvalue<0 then
	Response.write "<b>[操作失败]</b><p>等级不能小于0！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

sql="update 投票 set 等级='"&downvalue&"'"
set rs=conn.execute(sql)
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>更改投票最低经验值</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
<!--
.p9 {line-height: 150%; font-size: 9pt;}
.p12 {line-height: 150%; font-size: 12pt;}
body {line-height: 150%;font-size : 12pt;}
td {line-height: 150%;font-size : 12pt;}
A  {text-decoration: none;}
A:Hover  {text-decoration : none;}
a:visited {  color: #0000FF}
-->
</style>
</head>
<body bgcolor="#FFFFEC">
<h1 align="center"><font size="3" color="#0000FF">【更改投票最低经验值】</font></h1>
<hr noshade size="1" color=009900>
<p><b>[操作完成]</b></p>
<blockquote>
<p>已经将投票最低经验值更改为：<font color=FF0000><%=downvalue%></font><a href="javascript:history.go(-2)">【返回】</a></p>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>
