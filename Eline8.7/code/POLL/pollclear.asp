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
%><html>
<head>
<title>重新投票♀wWw.51eline.com♀</title>
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
    <td width="100%" height="1"><font size="2">您的位置&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font>&gt;&gt;<a href="pollclear.asp">清除得票数</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">【重新投票】</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[警告]</font></b></p>
<p>此操作将清除所有投票结果，您确定要这样做吗？ <a href="pollclearok.asp">是(YES)</a>　　<a href="javascript:history.go(-1)">否(NO)</a></p>
<br>
<hr noshade size="1" color=009900>
</body>
</html>
