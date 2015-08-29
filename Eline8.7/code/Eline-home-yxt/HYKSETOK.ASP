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
conn.open Application("sjjh_hykmdb")
rs.open "SELECT * FROM sethyk WHERE name='set'",conn
phyon=LCase(request.form("phyon"))
dhyon=LCase(request.form("dhyon"))
mon=LCase(request.form("mon"))
jon=LCase(request.form("jon"))
if phyon<>"on" then 
phyon="false" 
else 
phyon="true"
end if
if dhyon<>"on" then 
dhyon="false" 
else 
dhyon="true"
end if
if mon<>"on" then 
mon="false" 
else 
mon="true"
end if
if jon<>"on" then 
jon="false" 
else 
jon="true"
end if
conn.execute "update sethyk set phyt="&int(abs(Request.form("phyt")))&",phyf="&int(abs(Request.form("phyf")))&",dhy1t="&int(abs(Request.form("dhy1t")))&",dhy1f="&int(abs(Request.form("dhy1f")))&",dhy2t="&int(abs(Request.form("dhy2t")))&", dhy2f="&int(abs(Request.form("dhy2f")))&", dhy3t="&int(abs(Request.form("dhy3t")))&", dhy3f="&int(abs(Request.form("dhy3f")))&", dhy4t="&int(abs(Request.form("dhy4t")))&", dhy4f="&int(abs(Request.form("dhy4f")))&",j="&int(abs(Request.form("j")))&",m="&int(abs(Request.form("m")))&",mon="&mon&",dhyon="&dhyon&",phyon="&phyon&",jon="&jon&" where name='set'"
'j="&int(abs(Request.form("j")))&" 
%>
<html>
<head>
<title>会员数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理-高级管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
  <table border="1" width="35%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="209" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="32" colspan="4">
    高级设定<%=LCase(request.form("phyon"))%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="23%" align="center" nowrap height="27">
    泡点会员</td>
    <td bgcolor="#85C2E0"   width="4%" align="center" nowrap height="27">
   <%if rs("phyon") then%>开启<%else%>关闭<%end if%></td>
    <td bgcolor="#85C2E0"   width="21%" align="center" nowrap height="27">
    等级会员</td>
    <td bgcolor="#85C2E0"   width="5%" align="center" nowrap height="27">
   <%if rs("dhyon") then%>开启<%else%>关闭<%end if%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="23%" align="center" nowrap height="28">
    卖银子</td>
    <td bgcolor="#85C2E0"   width="4%" align="center" nowrap height="28">
  <%if rs("mon") then%>开启<%else%>关闭<%end if%></td>
    <td bgcolor="#85C2E0"   width="21%" align="center" nowrap height="28">
    卖金卡</td>
    <td bgcolor="#85C2E0"   width="5%" align="center" nowrap height="28">
 <%if rs("jon") then%>开启<%else%>关闭<%end if%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="26%" align="center" nowrap height="19" colspan="2">
    　</td>
    <td bgcolor="#85C2E0"   width="25%" align="center" nowrap height="19" colspan="2">
    　</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    泡点100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("phyf")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    泡点一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("phyt")%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    一级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("dhy1f")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    一级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("dhy1t")%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    二级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("dhy2f")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    二级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("dhy2t")%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    三级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("dhy3f")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    三级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("dhy3t")%></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    四级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("dhy4f")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    四级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("dhy4t")%></td>
  </tr>
  <tr>
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    1000万银子</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
<%=rs("m")%></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    1元金卡</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
<%=rs("j")%></td>
  </tr>
  </table>
<p align="center">更新完毕</p>
<p align="center">　</p>
<p align="center">本程序由阿男编写</p>
</body>
</html>
<%
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>