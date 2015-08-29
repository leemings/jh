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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 金,木,水,火,土,金币 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
ji=rs("金")
mu=rs("木")
shui=rs("水")
huo=rs("火")
tu=rs("土")
taiqiudian=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>属性转化♀wWw.51eline.com♀</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<div align="center"><b><font color="#008000">欢迎<%=sjjh_name%>光临属性转换处！</font></b></div>
<p align="center"><font color="#008000">在这里你将自己练的属性转换成金卡</font></p>
<p align="center"><font color="#008000">这样你就可以去买自己想要的卡片了! 2000点属性/1元!</font></p> 
<table border="1" align="center" bordercolor="#FF00FF">
  <tr> 
            <td valign="top" height="8" width="391" bgcolor="#008080" > 
              <div align="center">
<%if ji>2000 then%>
  <form method="POST" action="520.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#FFFFFF">金属性：<%=ji%>点，现金转换金卡<%=ji/2000%>元<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9"> 
    </font>
  </div>
</form>
<font color="#FFFFFF">
<%else%>
金属性少于2000不能转换 
<%end if%></font></div>
            </td>
          </tr>
          <tr> 
            <td valign="top" height="8" width="391" bgcolor="#008080" > 
              <div align="center">
<%if mu>2000 then%>
  <form method="POST" action="5201.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#FFFFFF">木属性：<%=mu%>点，现金转换金卡<%=mu/2000%>元<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9"> 
    </font>
  </div>
</form>
<font color="#FFFFFF">
<%else%>
木属性少于2000不能转换 
<%end if%>
</font>
              </div>
          <tr> 
            <td valign="top" height="8" bgcolor="#008080" > 
              <tr> 
            <td valign="top" height="8" bgcolor="#008080" > 
              <tr> 
            <td valign="top" height="8" bgcolor="#008080" > 
              <div align="center">
<%if shui>2000 then%>
  <form method="POST" action="5202.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#FFFFFF">水属性：<%=shui%>点，现金转换金卡<%=shui/2000%>元<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9"> 
    </font>
  </div>
</form>
<font color="#FFFFFF">
<%else%>
水属性少于2000不能转换 
<%end if%>
</font>
              </div>
          <tr> 
                                 <td valign="top" height="8" bgcolor="#008080" > 
              <tr> 
                                 <td valign="top" height="8" > 
              <tr> 
                                 <td valign="top" height="8" > 
              <div align="center">
<%if huo>2000 then%>
  <form method="POST" action="5203.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center">火属性：<%=huo%>点，现金转换金卡<%=huo/2000%>元<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9"> 
  </div>
</form>
<%else%>
火属性少于2000不能转换 
<%end if%>
              </div>
          <tr> 
                        <td valign="top" height="8" > 
              <tr> 
                        <td valign="top" height="8" > 
              <tr> 
                        <td valign="top" height="8" > 
              <tr> 
                        <td valign="top" height="8" > 
              <div align="center" style="width: 391; height: 178">
<%if tu>2000 then%>
  <form method="POST" action="5204.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center">土属性：<%=tu%>点，现金转换金卡<%=tu/2000%>元<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9"> 
  </div>
</form>
<%else%>
土属性少于2000不能转换 
<%end if%>
              </div>
  <tr>
  &nbsp;</tr>
</table>
              </body>
