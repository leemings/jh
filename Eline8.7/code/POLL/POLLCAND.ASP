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
if sjjh_grade<>10 or sjjh_name<>application("sjjh_user") then
	Response.Write "<b>[操作失败]</b><p>您没有管理候选人的权限！"
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
sql="Select count(*) As 数量 from 候选人"
set tmprs=conn.execute(sql)
hxr=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="select * from 候选人 order by 得票数 desc"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>候选人管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
<!--
.p9 {line-height: 150%; font-size: 9pt;}
.p12 {line-height: 150%; font-size: 12pt;}
body {line-height: 150%;font-size : 12pt;}
td {line-height: 135%;font-size : 10.5pt;}
A  {text-decoration: none;}
A:Hover  {text-decoration : none;}
a:visited {  color: #0000FF}
-->
</style>
<script language="JavaScript">
function del(name){
document.forms[1].delname.value=name
document.forms[1].delname.focus()
}
function load()
{var name=navigator.appName
var vers=navigator.appVersion
if(name=="Netscape")
{window.location.reload()
}else
{history.go(0)
}}</script>
</head>
<body bgcolor="#FFFFEC" topmargin="0" leftmargin="0">
<table border="1" cellpadding="2" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
<td width="100%" height="1"><font size="2">您的位置&gt;&gt;<a href="../jh.asp">江湖大街</a>&gt;&gt;<font color="#0000FF"><a href="poll.asp">江湖投票</a></font><font color="#000000">&gt;&gt;<a href="pollcand.asp">候选人管理</a></font></font></td>
</tr>
</table>
<h1 align="center"><font color="#0000FF" size="3">【候选人管理】</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[功能说明]</font></b></p>
<blockquote>
<p><font size="2">添加或删除候选人。（共有候选人：<font color=red><%=hxr%></font> 位；已投票：<font color=red><%=ytrs%></font>
票）<a href="javascript:load()">【刷新】</a><a href="javascript:history.go(-1)">【返回】</a></font></p>
<table border="1" align="center" cellpadding="4" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="#E0E0E0" width="80%">
<tr bgcolor="#0099FF" align="center">
<td><font color="#FFFFFF" size="2">序号</font></td>
<td><font color="#FFFFFF" size="2">候选人</font></td>
<td><font color="#FFFFFF" size="2">得票数</font></td>
<td><font color="#FFFFFF" size="2">得票率</font></td>
<td><font color="#FFFFFF" size="2">删除</font></td>
</tr>
<%
do while not(rs.eof or rs.bof)%>
<tr align="center">
<td><font size="2"><%=rs("id")%></font></td>
<td><font size="2"><%=rs("姓名")%></font></td>
<td><font size="2"><%=rs("得票数")%></font></td>
<td><font size="2">
<%if ytrs=0 then
	n=0
else
	n=int(rs("得票数")/ytrs*100)
end if
if n=0 then
	Response.Write "0%"
else
	Response.Write n&"%"
end if%>
</font></td>
<td><a href=javascript:del("<%=rs("姓名")%>")><font size="2">删除</font></a></td>
</tr>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</table>
</blockquote>
<hr noshade size="1" color=009900>
<form method="post" action="pollcandadd.asp">
<div align="center"> <font size="2"> 添加候选人：
<input type="text" name="addname" maxlength="10" size="10" style="font-size:12pt">
<input type="submit" name="add" value="添加" style="font-size:12pt">
<input type="reset" value="重写" style="font-size:12pt">
</font>
</div>
</form>
<form method="post" action="pollcanddel.asp">
<div align="center"> <font size="2"> 删除候选人：
<input type="text" name="delname" maxlength="10" size="10" style="font-size:12pt">
<input type="submit" name="del" value="删除" style="font-size:12pt">
<input type="reset" value="重写" style="font-size:12pt">
</font>
</div>
</form>
<hr noshade size="1" color=009900>
</body>
</html>