<%@ LANGUAGE=VBScript codepage ="936" %>
<%

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>

<head>
<title>赌场</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="main.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<table width="95%" border="0" cellspacing="0" cellpadding="0" bgcolor="#669900">
<tr>
<td width="22" height="21" align="left" valign="top"><img
src="jhimg/du1.gif" width="22" height="21"></td>
<td rowspan="3" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="67">
<div align="center"> <img src="jhimg/lswjs.gif" width="244" height="101">
</div>

</td>
</tr>
<tr>
<td height="50"><font color="#000000">&nbsp;<font size="2">黄头发蓝眼睛的洋人也看好<%=Application("aqjh_chatroomname")%>这块风水宝地，投资几亿开了这个大赌场。给国人争光，快去赢他一大票！：）</font></font></td>
</tr>
<tr>
<td>
<table border="0" cellpadding="0" cellspacing="0" width="433">
<!-- fwtable fwsrc="Untitled" fwbase="21.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
<tr>
<td><img src="image/spacer.gif" width="118" height="1"
border="0"></td>
<td><img src="image/spacer.gif" width="68" height="1" border="0"></td>
<td><img src="image/spacer.gif" width="223" height="1"
border="0"></td>
<td><img src="image/spacer.gif" width="24" height="1" border="0"></td>
<td><img src="image/spacer.gif" width="1" height="1" border="0"></td>
</tr>
<tr>
<td colspan="4">&nbsp;</td>
<td><img src="image/spacer.gif" width="1" height="10" border="0"></td>
</tr>
<tr>
<td rowspan="3"><a href="init21.asp"><img name="n21_r2_c1"
src="jhimg/21_r2_c1.gif" width="118" height="101" border="0"></a></td>
<td colspan="3">&nbsp;</td>
<td><img src="image/spacer.gif" width="1" height="59" border="0"></td>
</tr>
<tr>
<td rowspan="2"><a href="init21.asp"><img name="n21_r3_c2"
src="jhimg/21_r3_c2.gif" width="68" height="42" border="0"></a></td>
<td colspan="2">&nbsp;</td>
<td><img src="image/spacer.gif" width="1" height="19" border="0"></td>
</tr>
<tr>
<td background="jhimg/21bg.gif">　</td>
<td><img name="n21_r4_c4" src="jhimg/21_r4_c4.gif" width="24"
height="23" border="0"></td>
<td><img src="image/spacer.gif" width="1" height="23" border="0"></td>
</tr>
<tr>
<td colspan="4">&nbsp;</td>
<td><img src="image/spacer.gif" width="1" height="8" border="0"></td>
</tr>
</table>
</td>
</tr>
<tr>
<td></td>
</tr>
</table>
</td>
<td width="27" height="21" align="right" valign="top"><img
src="jhimg/du2.gif" width="27" height="21"></td>
</tr>
<tr>
<td height="249" width="4%" bgcolor="#669900">　</td>
<td height="249" bgcolor="#669900" width="4%">　</td>
</tr>
<tr>
<td width="22" height="2" align="left" valign="bottom"><img
src="jhimg/du3.gif" width="22" height="23"></td>
<td width="27" height="2" align="right" valign="bottom"><img
src="jhimg/du4.gif" width="27" height="23"></td>
</tr>
</table>

</body>

</html>