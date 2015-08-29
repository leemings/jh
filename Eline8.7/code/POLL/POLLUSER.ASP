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
if sjjh_grade<9 then
	Response.Write "<b>[操作失败]</b><p>你没有查阅已投票者的权限！"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select count(*) As 数量 from 投票者"
set tmprs=conn.execute(sql)
ytrs=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="select 姓名 from 投票者"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>查阅已投票者♀wWw.51eline.com♀</title>
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
    <td width="100%" height="1"><font size="2">您的位置&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font>&gt;&gt;<a href="polluser.asp">查阅已投票者</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">【查阅已投票者】</font></h1>
<hr noshade size="1" color=009900>
<blockquote>
<p><b><font color="#FF0000">[功能说明]</font></b></p>
<p><font size="2">下列用户已经投票，共 <font color=red><%=ytrs%></font> 人。<a href="javascript:history.go(-1)">【返回】</a></font></p>
</blockquote>
<table width="95%" border="1" align="center" cellspacing="0" cellpadding="4" bordercolorlight="#FF9999" bordercolordark="#FFFFFF">
<tr>
<td><font size="2">
<%do while not(rs.eof or rs.bof)%>
<%=rs("姓名")%>&nbsp;&nbsp;
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</font></td>
</tr>
</table>
<hr noshade size="1" color=009900>
</body>
</html>
