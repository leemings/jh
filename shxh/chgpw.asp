<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="<%=Application("Ba_jxqy_backgroundcolor")%>" text="#000000" background="<%=Application("Ba_jxqy_Backgroundimage")%>">
<p align="center"> <font color="FF0000" size="4">更改密码</font></p>
<form action="chgpw2.asp" method="post" id="form1" name="form1">
<table width="166" bgcolor="#CCCCCC" align="center" border="3">
<tr><td width="60">帐&nbsp;&nbsp;&nbsp;&nbsp;号</td><td width="118"><input name="account" maxlength="14" size="14" class="norsty"></td></tr>
<tr><td>姓&nbsp;&nbsp;&nbsp;&nbsp;名</td><td width="118"><input name="username" maxlength="14" size="14" class="norsty"></td></tr>
<tr><td>旧 密 码</td><td width="118"><input type="password" name="password" maxlength="14" size="14" class="norsty"></td></tr>
<tr><td>新 密 码</td><td width="118"><input type="password" name="password1" maxlength="14" size="14" class="norsty"></td></tr>
<tr><td>密码确认</td><td width="118"><input type="password" name="password2" maxlength="14" size="14" class="norsty"></td></tr>
<tr><td align="CENTER" colspan="2"><input type="submit" value=" 更 改 " class="norsty" id="submit1" name="submit1"> <input type="button" onclick="javascript:top.window.close();" value=" 关 闭 " class="norsty" id="button1" name="button1"></td></tr> 
</table> 
</form> 
</body></html> 
