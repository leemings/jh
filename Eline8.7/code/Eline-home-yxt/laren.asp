<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
larenseek=Request.Form("larenseek")
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>会员查询程序</title>
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
感谢你朋友这些人是你拉到我们江湖的！<br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">江湖拉人记录</font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="470" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
sql="SELECT * FROM 用户 where 等级>1 and 介绍人='"& larenseek &"' order by lasttime"
Set Rs=conn.Execute(sql)
%>
<td width="28" height="10">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="47" height="10">
<div align="center"><font color="#FFFFFF">姓名</font></div>
</td>
<td width="25" height="10">
<div align="center"><font color="#FFFFFF">性别</font></div>
</td>
<td width="63" height="10">
<div align="center"><font color="#FFFFFF">门派</font></div>
</td>
<td width="86" height="10">
<div align="center"><font color="#FFFFFF">身份</font></div>
</td>
<td width="75" height="10">
<div align="center"><font color="#FFFFFF">最后登陆时间</font></div>
</td>
<td width="51" height="10">
<div align="center"><font color="#FFFFFF">江湖等级</font></div>
</td>
<td width="35" height="10">
<div align="center"><font color="#FFFFFF">登录</font></div>
</td>
</tr>
<%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
<tr>
<td width="28" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="showuser.asp?username=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></font></a></div>
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
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("lasttime")%></font></div>
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
conn.close
%>
</table>
</td>
</tr>
<tr align="center">
    <td background="../jhmp/top3.gif" width="500" height="2"><font color="#FFFFFF"><%=Application("sjjh_chatroomname")%>程序修改</font></td>
</tr>
</table>
<div align="center"><font color="#000000">拉人总数:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font>
</div>
</body>

</html>