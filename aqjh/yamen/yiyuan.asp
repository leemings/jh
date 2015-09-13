<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>-快乐医院</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="700" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="318"><font size="4">&nbsp;</font></td>
<td width="460"><font color="#FF3300" size="4"><b>欢迎光临快乐医院</b></font> </td>
</tr>
</table>
<br>
<table width="700" border="0" cellpadding="0" cellspacing="0">
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
<tr bgcolor="#990000">
<td><img src="../images/lf2.gif"><img src="../images/lf3.gif" height="80"></td>
</tr>
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="700">
<tr>
<td>
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
align="center" >
<tr>
<td valign="top" width="100%" height="110">
<div align="center"> <br>
<table border="1" cellspacing="1" cellpadding="0" width="100%"
bordercolor="#000000" align="center">
<tr>
<td width="15%" nowrap>
<div align="center"> 花柳病患者 </div>
<td width="15%" nowrap>
<p align="center">梅毒病患者
</td>
<td width="11%">
<div align="center"> 尿道炎病患者 </div>
</td>
<td width="11%">
<div align="center"> 僵尸王病患者 </div>
</td>
</tr>
<tr>
<td width="15%" height="31">
<p align="center"><a href="yiyuan1.asp"><font color="#FF0000">请进</font></a>
<td width="15%" height="31">
  <p align="center"><a href="yiyuan2.asp"><font color="#FF0000">请进</font></a></td>
<td width="11%" nowrap height="31">
  <p align="center"><a href="yiyuan3.asp"><font color="#FF0000">请进</font></a></td>
<td width="11%" nowrap height="31">
  <p align="center"><a href="yiyuan6.asp"><font color="#FF0000">请进</font></a></td>
</tr>

</table>
<p>以后小心注意身体！ ！：（ </p> 
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<div align="center">
</div></body></html>