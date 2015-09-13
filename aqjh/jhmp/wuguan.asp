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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 师傅 from 用户 where 姓名='"&aqjh_name&"'",conn
sf=rs("师傅")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>快乐江湖</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=aqjh_name%></font>欢迎光临习武场</b></font></div>
<form method=POST action='wuguanok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
            <td align=center><font color=blue>师傅:<%=sf%></font>
            <%if trim(sf)<>"无" then%>
              <select name=money size=1>
                <option value="1000" selected> 打马步</option>
                <option value="10000">少林硬功</option>
                <option value="100000">峨眉心法</option>
                <option value="1000000">逍遥神剑</option>
                <option value="2000000">苗家蛊毒</option>
                <option value="10000000">浩然正气</option>
              </select>
              <%end if%>
</td>
</tr>
<tr>
<td  align=center>
            <%if trim(sf)<>"无" then%>
              <input type=submit value=开始练功 name="submit">
              <%end if%>
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                操作简介</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
            <%if trim(sf)="无" then%><div align="center">
            <font color=red>没有师傅不能操作！</font><BR></div>
            <%end if%>师傅领你去练功，选择你适合自己的心法(价钱是不同的哟！)才可以修天到至高无限的武功！</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><font color="#0000FF"><FONT color=#0000ff>&copy; 版权所有 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>快乐江湖</FONT></A></font></b></font>
</div>
</body>
</html>