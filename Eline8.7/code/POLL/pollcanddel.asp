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
	Response.Write "<b>[操作失败]</b><p>您没有删除候选人的权限！"
	Response.End
end if
delname=Server.HTMLEncode(Trim(Request.Form("delname")))
if delname="" then
	Response.Write "<b>[操作失败]</b><p>候选人名字为空，不能删除！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select 候选人 from 候选人 where 姓名='" & delname & "'",conn,3,3
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[操作失败]</b><p>候选人：<font color=red>" & delname & "</font> 不存在，不能删除！<a href=javascript:history.go(-1)>【返回】</a>"
	Response.End
end if
rs.close
conn.execute "delete * from 候选人 where 姓名='" & delname & "'"
set rs=nothing
conn.close
set conn=nothing
%><html>
<head>
<title>删除候选人</title>
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
<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
    <td width="100%" height="1"><font size="2">您的位置&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font><font color="#000000">&gt;&gt;<a href="pollcand.asp">候选人管理</a>&gt;&gt;</font><font color="#0000FF">删除候选人</font></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">【删除候选人】</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[操作完成]</font></b></p>
<blockquote>
<blockquote>
<p><font size="2">已经删除候选人：<font color="#FF0000"><%=delname%></font> ！<a href="javascript:history.go(-1)">【返回】</a>
</font>
</p>
</blockquote>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>