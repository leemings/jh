<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
usergrade=session("yx8_mhjh_usergrade")
if my="" then Response.Redirect "../error.asp?id=016"
%>
<HTML><HEAD><TITLE>通缉令</TITLE><LINK href="../style.css" rel=stylesheet>
</HEAD>
<BODY background='../chatroom/bg1.gif' leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
<TABLE border=0 cellPadding=0 cellSpacing=0 width=552 align="center">
<TBODY>
<TR>
<TD width="11">　</TD>
<TD align=middle vAlign=top width=661>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="601" height="123">
<TBODY>
<TR>
<TD width="11" height="103">　</TD>
<TD class=bg1
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" vAlign=top
width=557 height="103">
<div align="center">　
</div>
<div align="center"><b><font size="4" face="隶书">&nbsp;官府通缉令</font></b><font color="#FFFF00"><br>
</font>
</div>
<table border="1" cellspacing="0" cellpadding="0" width="100%" height="1%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<center>
<td  align="center">
<font  >被通缉者</font>
</td>
<td  align="center">
<font  >杀 人</font>
</td>
<td  align="center">
<font  >门 派</font>
</td>
<td  align="center">
<font  >身 份</font>
</td>
<td  align="center">
<font > 赏 金</font>
</td>
<td  align="center">
<font  >猎杀者</font>
</td>
<td  align="center">
<font  >领 赏</font>
</td>
</center>
</tr>
<%
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 等级<101 and 杀人>=10 order by 杀人 desc"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td height="1">
<center>
<font ><%=rs("姓名")%></font>
</center>
</td>
<td height="1">
<center>
<font ><%=rs("杀人")%></font>
</center>
</td>
<td height="1">
<center>
<font ><%=rs("门派")%></font>
</center>
</td>
<td height="1">
<center>
<font ><%=rs("身份")%></font>
</center>
</td>
<td height="1">
<center>
<font ><%=rs("杀人")*50000%>两</font>
</center>
</td>
<td height="1">
<center>
<font >
<% if rs("杀人者")<>"" then%>
<%=rs("杀人者")%></font>
<%else%>
等待中
<%end if%>
</center>
</td>
<td height="1">
<center>
<font  >
<%if my=rs("杀人者") and my<>rs("姓名") then%>
</font><a href="lingshang.asp?name=<%=rs("姓名")%>">领赏
</a>
<font  >
<%end if%>
<%if usercorp="官府" and usergrade>=lockipright then%>
</font>
<a href='zhanshou.asp?name=<%=rs("姓名")%> '>斩首
</a>
<font  color="#FFFF00">
</font>
<font  >
<%end if%>
操作
</font>
</center>
</td>
</tr>
<%
rs.movenext
loop  
%>
</table>
</TD>
<TD vAlign=top width="19" height="103">　</TD></TR>
<TR>
<TD width="19" height="20">　</TD></TR>
</TBODY></TABLE>
</TD>
<TD width="13">　</TD></TR>
</TBODY></TABLE>
<p align="center">说明：杀人超过10个就会被官府通缉，大侠们可以猎杀这些“被通缉者”，然后来领重赏！</p>
</BODY></HTML>
<%
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing       
%>