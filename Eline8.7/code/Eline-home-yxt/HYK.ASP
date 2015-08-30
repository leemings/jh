<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>会员数据库管理♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<table border="1" width="14%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="87" align="center">
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="51%" align="center" nowrap height="31">
    <a href="addhyk.asp">添加会员卡</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="51%" align="center" nowrap height="31">
    <a href="finehyk.asp">查找会员卡</a></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="51%" align="center" nowrap height="31">
    <a href="hykset.asp">会员卡管理</a></td>
  </tr>
</table>
<p align="center">　</p>
<p align="center">本程序由酷帅编写</p>
</body>
</html>