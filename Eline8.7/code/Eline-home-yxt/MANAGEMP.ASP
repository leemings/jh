<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
id=Request.QueryString("ID")
sql="Select * from p where a='"&Id&"'"
rs=conn.execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','修改门派资料：["&id&"]')"
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 宋体}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#FAF0E2" text="#000000" link="#000080" alink="#800000" vlink="#000080">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="498" height="31">
<tr align="center" bgcolor="#336633">
<td width="100%" height="25"><b><font color="#FFFFFF" size="4">门派内容修改</font></b></td>
</tr>
</table>
</div>
<form action="updatemp.asp?subject=<%=rs("a")%>" method=POST>
  <ul>
    <table border=1 cellspacing=0 cellpadding=3 align="center" width="500" bordercolordark="#FFFFFF">
      <tr> 
        <td bgcolor="#F7E7D0" width="180"><font size="2" class="c">门派</font></td>
        <td width="308"> 
          <input name="mp" size=40 maxlength=50 value="<%=RS("a")%>">
          <input name="id" type="hidden" size=40 maxlength=50 value="<%=RS("id")%>">
        </td>
      </tr>
      <tr> 
        <td bgcolor="#FAF0E2" width="180"><font class="c" size="2">掌门<font color="#cc0000">&nbsp;</font></font></td>
        <td width="308"> 
          <input name="zm" value="<%=RS("b")%>" size=40 maxlength=50>
        </td>
      </tr>
      <tr> 
        <td bgcolor="#F7E7D0" width="180"><font class="c" size="2">简介</font></td>
        <td width="308"> 
          <input name="sm" value="<%=RS("d")%>" size=40 maxlength=50>
        </td>
      </tr>
      <tr> 
        <td width="180" ><font class="c" size="2">限制说明<font color="#cc0000"></font></font></td>
        <td width="308"><font size="3" class="c" color="#000000"> 
          <input type="text" name="xzsm" size="40" value="<%=RS("e")%>">
          </font> </td>
      </tr>
      <tr> 
        <td valign=top width="180" height="32" bgcolor="#F7E7D0"><font size="2">表达式</font></td>
        <td width="308" height="32"> 
          <input name="bds" value="<%=RS("g")%>" size=40 maxlength=200>
        </td>
      </tr>
      <tr> 
        <td valign=top width="180"><font size="2">门派基金</font></td>
        <td width="308"> 
          <input name="jj" value="<%=RS("h")%>" size=40 maxlength=200>
        </td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c" color="#000000"><br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="更新">
        </font></p>
    </div>
  </ul>
</form>
</body>
</html>
<%
set rs=nothing
%>