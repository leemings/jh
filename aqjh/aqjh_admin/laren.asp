<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
larenseek=Request.Form("larenseek")
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>拉人查询程序</title>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
<br>
感谢你朋友这些人是你拉到我们江湖的！<br>
<table width="650" border="1" cellpadding="0" cellspacing="0" height="13" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM 用户 where 等级>1 and 介绍人='"& larenseek &"' order by lasttime"
Set Rs=conn.Execute(sql)
%>
<td width="20" height="10">
<div align="center">ID</font></div>
</td>
<td width="47" height="10">
<div align="center">姓名</font></div>
</td>
<td width="25" height="10">
<div align="center">性别</font></div>
</td>
<td width="63" height="10">
<div align="center">门派</font></div>
</td>
<td width="75" height="10">
<div align="center">身份</font></div>
</td><td width="75" height="10">
<div align="center">注册ip</font></div>
</td><td width="75" height="10">
<div align="center">最后ip</font></div>
</td>
<td width="75" height="10">
<div align="center">最后登陆时间</font></div>
</td>
<td width="51" height="10">
<div align="center">江湖等级</font></div>
</td>
<td width="35" height="10">
<div align="center">登录</font></div>
</td>
</tr>
<%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
<tr>
<td width="28" height="30">
<div align="center"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="showuser.asp?username=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><%=rs("性别")%></font></div>
</td>
<td width="63" height="30">
<div align="center"><%=rs("门派")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><%=rs("身份")%></font></div>
</td><td width="75" height="30">
<div align="center"><%=rs("regip")%></font></div>
</td><td width="75" height="30">
<div align="center"><%=rs("lastip")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><%=rs("lasttime")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><%=rs("等级")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><%=rs("times")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
conn.close
%>
</table>
<div align="center"><font color="#000000">拉人总数:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font>
</div>
</body></html>