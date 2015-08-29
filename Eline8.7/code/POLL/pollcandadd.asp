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
	Response.Write "<b>[操作失败]</b><p>您没有添加候选人的权限！"
	Response.End
end if
addname=Server.HTMLEncode(Trim(Request.Form("addname")))
'addname=Trim(Request.Form("addname"))
if addname="" then
	Response.Write "<b>[操作失败]</b><p>候选人名字为空，不能添加！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,姓名 from 用户 where 姓名='" & addname & "'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[操作失败]</b><p>江湖中无此人，不能添加！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.End
end if
id=rs("ID")
rs.close
rs.open "Select 候选人 from 候选人 where 候选人="&id,conn,3,3
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[操作失败]</b><p>此人已经是候选人了，不能重复添加！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.End
end if
rs.close
conn.execute "insert into 候选人(候选人,姓名,得票数) values ('" & id & "','" & addname & "',0)"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>添加候选人</title>
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
<body bgcolor="#FFFFEC" topmargin="0" leftmargin="0">
<table border="1" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1" cellpadding="0">
<tr>
    <td width="100%" height="1"><font size="2">您的位置&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font><font color="#000000">&gt;&gt;<a href="pollcand.asp">候选人管理</a>&gt;&gt;</font><font color="#0000FF">添加候选人</font></font></td>
</tr>
</table>
<h1 align="center"><font color="#0000FF" size="3">【添加候选人】</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[操作完成]</font></b></p>
<blockquote>
<blockquote>
<p><font size="2">已经添加候选人：<font color="#FF0000"><%=addname%></font> ！<a href="javascript:history.go(-1)">【返回】</a>
</font>
</p>
</blockquote>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>