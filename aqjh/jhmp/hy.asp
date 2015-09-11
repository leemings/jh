<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"


%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>会员查询程序</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="../jhimg/bk_hc3w.gif">
<p align="center"> <font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
<br>
感谢这些朋友对我们江湖的大力支持！加入会员点<a href="../chat/hy.asp" target="_blank">这里</a><br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">江 湖 会 员</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='正常' and 会员等级>0 order by 会员日期 ",conn
%>
<td width="28" height="18">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="47" height="18">
<div align="center"><font color="#FFFFFF">姓名</font></div>
</td>
<td width="25" height="18">
<div align="center"><font color="#FFFFFF">性别</font></div>
</td>
<td width="63" height="18">
<div align="center"><font color="#FFFFFF">门派</font></div>
</td>
<td width="86" height="18">
<div align="center"><font color="#FFFFFF">身份</font></div>
</td>
<td width="40" height="18">
<div align="center"><font color="#FFFFFF">会员级</font></div>
</td>
<td width="75" height="18">
<div align="center"><font color="#FFFFFF">会员结束日期</font></div>
</td>
<td width="51" height="18">
<div align="center"><font color="#FFFFFF">江湖等级</font></div>
</td>
<td width="35" height="18">
<div align="center"><font color="#FFFFFF">登录</font></div>
</td>
</tr>
<%
dengji1=0
dengji2=0
dengji3=0
dengji4=0
dengji5=0
dengji6=0
dengji7=0
dengji8=0
do while not rs.bof and not rs.eof
Select Case rs("会员等级")
Case 1
	dengji1=dengji1+1
case 2
	dengji2=dengji2+1
case 3
	dengji3=dengji3+1
case 4
	dengji4=dengji4+1
case 5
	dengji5=dengji5+1
case 6
	dengji6=dengji6+1
case 7
	dengji7=dengji7+1
case 8
	dengji8=dengji8+1

end select
%>
<tr>
<td width="28" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("性别")%></font></div>
</td>
<td width="63" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("门派")%></font></div>
</td>
<td width="86" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("身份")%></font></div>
</td>
<td width="40" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("会员等级")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("会员日期")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("等级")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</td>
</tr>
<tr align="center">
<td background="top3.gif" width="500" height="2"><font color="#FFFFFF"><%=Application("aqjh_chatroomname")%>版权所有</font></td>
</tr>
</table>
<div align="center"><font color="#000000">一级用户：</font><b><font color="#0000FF"><%=dengji1%></font></b>
<font color="#000000">二级会员:</font><b><font color="#0000FF"><%=dengji2%></font></b><font color="#000000">人
三级会员：</font><font color="#0000FF"><b><%=dengji3%></b></font><font color="#000000">人
四级会员：</font><font color="#0000FF"><b><%=dengji4%></b></font><font color="#000000">人
五级会员：</font><font color="#0000FF"><b><%=dengji5%></b></font><font color="#000000">人
六级会员：</font><font color="#0000FF"><b><%=dengji6%></b></font><font color="#000000">人
七级会员：</font><font color="#0000FF"><b><%=dengji7%></b></font><font color="#000000">人
八级会员：</font><font color="#0000FF"><b><%=dengji8%></b></font><font color="#000000">人<br>
会员总数：</font><b><font color="#0000FF"><%=(dengji1+dengji2+dengji3+dengji4+dengji5+dengji6+dengji7+dengji8)%></font></b><font color="#000000">人</font>
</div>
</body>
</html>