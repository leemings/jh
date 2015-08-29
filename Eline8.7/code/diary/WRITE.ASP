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
%>
<html>
<head>
<title>开始写日记♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: 宋体 }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body bgcolor="#BABABA" topmargin="0" leftmargin="0" background="../bgcheetah.gif">
<form method="POST" action="writeok.asp">
  <table border="1" width="498" cellspacing="0" cellpadding="1" align="center" background="backgroup1.gif">
    <tr> 
      <td colspan="4" height="33"> 
        <div align="center"> <font size="2" class="c"><font size="3"><b><font color="#0000FF"><%=sjjh_name%></font></b><font size="2" color="#0000FF">大侠的江湖心情日记</font></font></font> 
        </div>
</td>
</tr>
    <tr> 
      <td height="6" width="220">时间: <font color="#0000FF" size="2"> <%=now()%> 
        </font></td>
      <td height="6" colspan="2">天气： 
        <input type="text"
name="tianqi" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="8">
</td>
      <td height="6" width="136"> 心情:<font color="#0000FF" size="2"> 
        <input type="text"
name="xinqing" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9">
</font></td>
</tr>
</table>
<table border="1" width="498" bgcolor="#33CCFF" cellspacing="0" cellpadding="1" align="center">
    <tr> 
      <td width="37" bgcolor="#3399FF">主题:</td>
      <td width="267"> 
        <input type="text"
name="title" size="50"
style="font-family: Tahoma; font-size: 12px"
maxlength="35">
      </td>
</tr>
</table>
  <table width="498" border="0" cellspacing="0" cellpadding="0" align="center" height="129" background="backgroup1.gif">
    <tr>
<td height="172" width="494" align="left" valign="top">
<div align="center"><font size="+1">日记内容</font><span class="txt">
<textarea name="diary" cols="80" style="font-family: Tahoma; font-size: 12px" rows="20">请在这里写上你的心里话!</textarea>
</span></div>
</td>
</tr>
</table>
<table border="1" width="498" bgcolor="#33CCFF" cellspacing="0" cellpadding="1" align="center">
<tr bgcolor="#cccccc">
<td bgcolor="#3399FF">&nbsp;</td>
</tr>
</table>
<div align="center"><br>
<input type="submit" value="我写完了OK" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</div>
</form>
<div align="center">
<br>
</div>
</body>

</html>
