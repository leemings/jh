<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>睡眠人员名单</title>
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
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="318">&nbsp;</td>
<td width="460"><font color="#FF3300"><b>睡眠人员名单</b></font> </td>
</tr>
</table>
<br>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td bgcolor="#000000">&nbsp;</td>
  </tr>
  <tr bgcolor="#990000"> 
    <td> 
      <div align="left"><img src="eline.jpg" width="778" height="80"></div></td>
  </tr>
  <tr> 
    <td bgcolor="#000000">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="778">
<tr>
<td>
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
align="center" >
<tr>
<td colspan="2" valign="top" width="100%" height="110">
<div align="center"> <br>
<table border="1" cellspacing="1" cellpadding="0" width="100%"
bordercolor="#000000" align="center">
<tr>
<td width="15%" nowrap>
<div align="center"> 睡眠 </div>
<td width="15%" nowrap>
<div align="center"> 门派 </div>
</td>
<td width="11%">
<div align="center"> 身份 </div>
</td>
<td width="29%">
<div align="center"> 请 求 </div>
</td>
<td width="21%">&nbsp;</td>
<td width="9%" nowrap>　</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='眠'",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="15%" height="31"><%=rs("姓名")%>
<td width="15%" height="31"><%=rs("门派")%></td>
<td width="11%" nowrap height="31"><%=rs("身份")%></td>
<%if rs("状态")="眠" then %>
<td width="29%" height="31">
<div align="center"> <a href="shifang1.asp?name=<%=rs("姓名")%>"><font color="#008000">交钱8000万两+法力5000点解救</font></a>  
<%if sjjh_grade=10 then%> 
 <a href="zhanshou.asp?name=<%=rs("姓名")%>"><font color="#FF0000"></font></a>  
<%end if%> <%if sjjh_grade>=6 then%>   
  |<a href="shuchu1.asp?id=<%=rs("id")%>"><font color="#FF0000">解救</font></a> 
<%end if%> </div>   
</td>
<% else
if rs("登录")>now() then%>
<td width="21%" height="31">
	<div align="center"> <font color="#FF0033"></font>
	</div>
	</td>
<%else%>
	<td width="9%" height="31">
	<div align="center"> <font color="#FF0033"></font>
	</div>
	</td>
	<% conn.execute("update 用户 set 状态='正常',登录=now() where 姓名='"&rs("姓名")&"'")
end if%>
<%end if%>
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
<p>以后小心点噢！ ！：（ </p> 
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>

