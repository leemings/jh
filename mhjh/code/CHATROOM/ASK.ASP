<%
username=session("yx8_mhjh_username")
if username="" then response.redirect "../error.asp?id=016"
if session("yx8_mhjh_usercorp")<>"官府"  then response.redirect "../error.asp?id=046"
%>
<html>
<head>
<title>答对有奖</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body background='bg1.gif'>
<form method="post" action="asking.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#33CCCC">
<td colspan="2" height="27" >快乐江湖抢答</td>
</tr>
<tr>
<td valign="top">问：</td>
<td >
<textarea name="ask" rows="3" wrap="VIRTUAL" cols="30"></textarea>
</td>
</tr>
<tr><td >答：</td><td><input type="text" name="reply" size="10" maxlength="10"></td></tr>
<tr>
<td>
奖：</td><td><input type=text name='silver'  maxlength=4 size=4>
</td>
</tr>
<tr>
<td colspan=2 align=center><input type="submit" name="Submit" value="提 交"> <input type="button"  value="关 闭" onclick='javascript:top.window.close();'></td>
</tr>
<tr>
<td colspan="2">
<div align="center">出题人：<%=username%> </div>
</td>
</tr>
</table>
</form>
<p align="center">网管负责出题，对于答对的聊友，系统将给予奖励！<br>
注意!不得利用此功能做假! </p>
</body>
</html>
