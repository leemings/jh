<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
%>
<HTML><HEAD><TITLE>通缉令</TITLE>
<LINK href="css.css" rel=stylesheet>
</HEAD>
<BODY bgColor=#FFDDF2  text="#FFFFFF" background="bg1.gif">
<div align="center">　
</div>
<div align="center"><b><font size="4" face="隶书" color="#800000">&nbsp;</font><font size="4" face="幼圆" color="#800000">官府通缉令</font></b><font color="#FFFF00"><br>
</font>
</div>
<div align="center">
  <center>
<table border="1" cellspacing="1" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#000000">
<tr bgcolor="#DFEDFD">
<center>
<td height="14" bgcolor="#005B00" align="center">
<font color="#FFFFFF" >被通缉者</font>
</td>
<td height="14" bgcolor="#005B00" align="center">
<font  color="#FFFFFF">杀 人</font>
</td>
<td height="14" bgcolor="#005B00" align="center">
<font color="#FFFFFF"> 赏 金</font>
</td>
</center>
</tr>
<%
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 杀人,姓名 FROM 用户 where 杀人>=10 order by 杀人 desc"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor="#f7f7f7">
<td bgcolor="#FFFFFF" height="1">
<font color="#000000"><%=rs("姓名")%></font>
</td>
<td bgcolor="#FFFFFF" height="1">
<font color="#000000"><%=rs("杀人")%></font>
</td>
<td bgcolor="#FFFFFF" height="1">
<font color="#000000"><%=rs("杀人")*20000%>两</font>
</td>
</tr>
<%
rs.movenext
loop  
%>
</table>
  </center>
</div>

<p align="center"><font color="#800000">杀死通缉者就可<br>以领取官府奖励</font></p>
</BODY></HTML>
<%
rs.Close 
set rs=nothing             
conn.Close           
set conn=nothing
%>