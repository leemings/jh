<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
chuwuseek=Request.Form("chuwuseek")
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>储物查询程序</title>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
<table width="650" border="1" cellpadding="0" cellspacing="0" height="13" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM w where a="&chuwuseek
Set Rs=conn.Execute(sql)
%>
<td height="10">
<div align="center">物品</font></div>
</td>
<td height="10">
<div align="center">数量</font></div>
</td>
</tr>
<%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
<tr>
<td height="20">
<div align="center"><font color="#FF9900"><%=rs("b")%></font></div>
</td>
<td height="20">
<div align="center"><%=rs("c")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
conn.close
%>
</table>
</body></html>