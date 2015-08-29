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
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<table border="1" width="34%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="132" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="132">
    <form method="POST" action='listhyk.asp?lx=3'>
      <p align="center">
      <select size="1" name="hykcx">
      <option selected>卡号</option>
      <option>所有者</option>
      </select><input type="text" name="hykcxname" size="20">
      <input type="submit" value="确定" name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid"></p>
    </form>
    <a href="listhyk.asp?lx=1">未使用的会员卡</a><p><a href="listhyk.asp?lx=2">会 员 卡 一 览</a></td>
  </tr>
  </table>
<p align="center">　</p>
<p align="center">本程序由阿男编写</p>
</body>
</html>