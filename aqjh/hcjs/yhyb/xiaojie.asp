<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 性别 FROM 用户 WHERE 姓名='" & aqjh_name & "'",conn
sex=rs("性别")
rs.close
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>人妖馆</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../../bg.gif">
<table border=1 bgcolor="#FFFFFF" align=center width=650 cellpadding="0" cellspacing="1" bordercolor="#FFCC00">
  <tr bgcolor="#669900"> 
<td height="16" colspan="6" width="533"> 
<p align=center class="p9"><font color="#FF3333"><b><font color="#FFFFFF">人妖馆</font></b></font></p>
<tr bgcolor=#74E76D bordercolor="#000000"> 
    <td width="45" height="15" align="center" bordercolor="#FF6600" bgcolor="#FFFFCC"> 
      <div align="center">
<font size="2" color="#FF0000">妓男</font></div>
</td>
    <td width="49" height="15" align="center" bordercolor="#FF6600" bgcolor="#FFFFCC"> 
      <div align="center">
<font size="2" color="#FF0000">美貌</font></div>
</td>
    <td width="93" height="15" align="center" bordercolor="#FF6600" bgcolor="#FFFFCC"> 
      <div align="center">
  <p align="center"><font size="2" color="#FF0000">和其聊天的价格</font></div>
</td>
    <td width="122" height="15" align="center" bordercolor="#FF6600" bgcolor="#FFFFCC"> 
      <div align="center">
  <p align="center"><font size="2" color="#FF0000">和其聊天增加的内力</font></div>
</td>
<td height="15" width="102" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">进入房间</font></div>
</td>
<td height="15" width="102" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">为他赎身</font></div>
</td>
</tr>
<!--#include file="jiu.asp"-->
<%
sql="SELECT id ,姓名,美貌度 FROM 妓男"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor=#DEAD63> 
    <td bgcolor="#FFFFFF" width="120"> 
      <center>
<font size="2"><%=rs("姓名")%><span class="p9"><font color="#6666FF"></font></span> 
</font> 
</center>
    <td bgcolor="#FFFFFF" width="80"> 
      <div align="center"><font size="2"><%=rs("美貌度")%></font> </div>
<td bgcolor="#FFFFFF" width="93"> 
<div align="center"><font size="2"><%=rs("美貌度")*5%> </font></div>
</td>
<td bgcolor="#FFFFFF" width="122"> 
<div align="center"><font size="2"><%=int(rs("美貌度")/2)%> </font></div>
</td>
<td bgcolor="#FFFFFF" width="102"> 
<center>
<font size="2"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">进入房间</span></a></font> 
</center>
</td>
<td>
<div align="center"><font size="2"><a href=shushen.asp?id=<%=rs("id")%>>赎出他来</a></font></div>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
connt.close
set connt=nothing
%>
</table>
<div align="center"><br>
  <a href="index.asp">返回</a></div><br><br>
<div align="center"><font color="#000000"><b>版权:『快乐江湖总站』</b></font></div>
</BODY></HTML>