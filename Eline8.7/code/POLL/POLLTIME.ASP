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
	Response.Write "<b>[操作失败]</b><p>你没有更改投票时间的权限！"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select  * from 投票"
set rs=conn.execute(sql)
ks=rs("开始时间")
jz=rs("截止时间")
rs.close
set rs=nothing
conn.close
set conn=nothing
%><html>
<head>
<title>更改投票起止时间</title>
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
    <td width="100%" height="1"><font size="2">您的位置&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font>&gt;&gt;<a href="polltime.asp">更改投票时间</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">【更改投票起止时间】</font></h1>
<hr noshade size="1" color=009900>
<blockquote>
<p><b><font color="#FF0000">[功能说明]</font></b></p>
<p><font size="2">本功能用于设置允许投票的时间段。格式为“年-月-日 时:分:秒”，例：“2000-02-17 08:00:00”（全部为半角字符）。</font></p>
</blockquote>
<table border="1" align="center" bgcolor="#E0E0E0" cellpadding="4" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0">
<form method="post" action="polltimeok.asp">
<tr>
<td><font size="2">开始投票时间：</font></td>
<td>
<font size="2">
<input type="text" name="begintime" maxlength="19" size="19" style="font-size:12pt" value="<%=ks%>">
</font>
</td>
</tr>
<tr>
<td><font size="2">终止投票时间：</font></td>
<td>
<font size="2">
<input type="text" name="endtime" maxlength="19" size="19" style="font-size:12pt" value="<%=jz%>">
</font>
</td>
</tr>
<tr>
<td align="right"><font size="2">选择操作：</font></td>
<td align="center">
<font size="2">
<input type="submit" name="Submit" value="更改" style="font-size:12pt">
<input type="button" value="放弃" style="font-size:12pt" onclick=javascript:history.go(-1)>
</font>
</td>
</tr>
</form>
</table>
<hr noshade size="1" color=009900>
</body>
</html>