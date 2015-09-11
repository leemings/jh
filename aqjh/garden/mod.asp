<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>〖花园大赛〗</TITLE>
<link rel="stylesheet" href="../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<!--#include file="data.asp"-->
<center>
<font color="red"><b>花园大赛报名处</b></font><BR><br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from 花园设置"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
%>
<font color="#0000FF">报名开始时间：<font color=red><%=rs("报名开始")%></font> <BR>报名结束时间：<font color=red><%=rs("报名结束")%></font></font>
<%if CDate(rs("报名开始"))>now() then
   	Response.Write "<br><br><b>[提示]</b>报名还没在开始呢？"
	Response.End
  elseif CDate(rs("报名结束"))<now() then
     	Response.Write "<br><br><b>[提示]</b>报名已经结束了，下一次再来吧！"
	Response.End
  else
%>
<table border=1 bgcolor="#669900" align=center width=350 cellpadding="10" cellspacing="13">
<tr>
<td bgcolor=#FFFFFF align="center"> 
<table bgcolor="#FFFFFF" bordercolor="#FFFF00" border="0">
<tr> 
<td bgcolor="#FFFFFF" height="39">&nbsp;</td>
<td bgcolor="#FFFFFF" height="39" align="center"> 
<form method=POST action='modify.asp'>
<table border="0" cellpadding="0">
<tr> 
<td align="center">参赛主题：</td>
<td align="center"> 
<input type=text name=hyname size=12 maxlength="11" value=""> <font size=2 color=red>想表现什么主题</font>
</td>
</tr>
<tr> 
<td align="center">主题说明：</td>
<td align="center"> 
<input type=text name=hyTitle size=26 maxlength="32" value="">
</td>
</tr>
</table><br>
<input type=submit value=报名 name="submit">
<input type=button value=关闭 onClick="window.close()" name="button">
</form>
</td>
</tr>
</table>
</table>
<%end if
rs.close
end if
%>
<font color=blue>主题与说明中不要使用空格！</font>
</center> 
</body></html>