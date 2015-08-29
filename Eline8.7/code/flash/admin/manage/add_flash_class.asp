<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<title>添加类别</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="../style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="bkg.gif">
<form method="post" action="save_flash_class.asp" name="class">
  <table width="500" border="0" cellspacing="1" cellpadding="0" align="center">
    <tr> 
      <td width="120" height="40" align="center">类别名称：</td>
      <td height="40"> 　 
        <input type="text" name="ClassName" size="50" class="button" maxlength="100">
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2" height="40"> 
        <input type="submit" name="Submit" value="提交" class="button">
        　　 
        <input type="reset" name="reset" value="重新填写" class="button">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
