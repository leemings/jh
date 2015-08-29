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
id=request.querystring("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from j where id="& id,conn
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
<form method="POST" action="cdiary.asp?id=<%=rs("id")%>">
<table border="1" width="498" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
<tr bgcolor="#6699FF">
<td colspan="4" height="33">
<div align="center"> <font size="2" class="c"><font size="3"><b><font color="#0000FF"><%=sjjh_name%></font></b><font size="2" color="#0000FF">大侠的江湖心情日记</font></font></font>
</div>
</td>
</tr>
<tr bgcolor="#3399FF">
<td height="6" width="220">时间:<font color="#FFFFFF"> <font size="2"> <%=rs("b")%></font></font></td>
<td height="6" colspan="2">天气：
<input type="text"
name="tianqi" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="8" value="<%=rs("d")%>">
</td>
<td height="6" width="136"> 心情:<font color="#0000FF" size="2">
<input type="text"
name="xinqing" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9" value="<%=rs("c")%>">
</font></td>
</tr>
</table>
<table border="1" width="498" bgcolor="#33CCFF" cellspacing="0" cellpadding="1" align="center">
<tr bgcolor="#cccccc">
<td width="37" bgcolor="#3399FF">主题:</td>
<td width="267" bgcolor="#3399FF">
<input type="text"
name="title" size="50"
style="font-family: Tahoma; font-size: 12px"
maxlength="35" value="<%=rs("e")%>">
</td>
</tr>
</table>
<table width="498" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#3399FF" height="129">
<tr>
<td height="172" width="494" align="left" valign="top">
<div align="center"><font size="+1">日记内容</font><span class="txt">
<textarea name="diary" cols="80" style="font-family: Tahoma; font-size: 12px" rows="20"><%=rs("f")%></textarea>
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
<input type="submit" value="　 修改　" name="B1" tyle="font-family: Tahoma; font-size: 12px">
<input type="button" name="ok" value="　返 回　" onClick=javascript:history.go(-1)>
<input type="button" name="ok2" value=" 关 闭　" onClick=javascript:window.close()>
</div>
</form>
<div align="center">
<br>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing

%>