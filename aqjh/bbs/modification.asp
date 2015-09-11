<!-- #include file="setup.asp" -->
<%top%>
<SCRIPT src="inc/birthday.js"></SCRIPT>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 忘记密码</td>
</tr>
</table><br>

<br>
<form action="RecoverPasswd.asp" method="post">

<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" class=a1 height="20"><div align="center">
忘记密码</div>
</td></tr><tr>
<td height="19" width="328" valign="top" class=a3>
<div align="center">用户名称：<input size="25" name="username" value="<%=Request.Cookies("username")%>"></div>
<div align="center">出生日期：<input onfocus="show_cele_date(birthday,'','',birthday)" name="birthday" size="25"><br>
<input type="submit" value=" 确定 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit">
</div></td></tr> </table></form><center><a href=javascript:history.back()>BACK </a><br>

<%htmlend%>