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
<title><%=Application("aqjh_chatroomname")%>国家金币排行</title>
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
<p align="center">
<%
id=request("id")
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">爱 情 国 家</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 国家='"&id&"' order by 金币 desc",conn
%>
<td width="60">
<div align="center"><font color="#FFFFFF">姓名</font></div>
</td>
<td width="31">
<div align="center"><font color="#FFFFFF">金币</font></div>
</td>
<td width="27">
<div align="center"><font color="#FFFFFF">性别</font></div>
</td>
<td width="74">
<div align="center"><font color="#FFFFFF">身份</font></div>
</td>
<td width="75">
<div align="center"><font color="#FFFFFF">智力数</font></div>
</td>
<td width="66">
<div align="center"><font color="#FFFFFF">知质数</font></div>
</td>
<td width="65">
<div align="center"><font color="#FFFFFF">重生次数</font></div>
</td>
<td width="32">
<div align="center"><font color="#FFFFFF">离派次数</font></div>
</td>
<td width="73">
<div align="center"><font color="#FFFFFF">掌门操作</font></div>
</td>
<td width="73">
<div align="center"><font color="#FFFFFF">掌门操作</font></div>
</td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr>
<td width="60" height="30">
<div align="center"><a href="../guojia/mt.asp?action=<%=rs("姓名")%>"><font color="#FFFFFF"><%=rs("姓名")%></font></a></div>
</td>
<td width="31" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("金币")%></font></div>
</td>
<td width="27" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("性别")%></font></div>
</td>
<td width="74" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("身份")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("智力")%></font></div>
</td>
<td width="66" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("知质")%></font></div>
</td>
<td width="65" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("重生")%></font></div>
</td>
<td width="32" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("离派")%></font></div>
</td>
<td width="73" height="30">
<div align="center"><a href="gj3.asp?you=<%=rs("姓名")%>&amp;id=<%=rs("国家")%>"><font color="#FFFFFF">逐出国境</font></a></div>
</td>
<td width="73" height="30">
<div align="center"><a href="gj8.asp?you=<%=rs("姓名")%>&id=<%=rs("国家")%>"><font color="#FFFFFF">取消册封</font></a></div>
</td>
</tr>
<%
rs.movenext
filename=filename+1 
if filename>20 then Exit Do 
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
<td background="top3.gif" width="500" height="28"><font color="#FFFFFF">版权所有『快乐江湖总站』</font></td>
</tr></table></body></html>