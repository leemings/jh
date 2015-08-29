<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_plus_check.asp"-->
<% 
dim nam
stats="头像集锦"
call nav() 
call head_var(2,0,"","")
if not founduser or membername="" then
	errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用这个功能。。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
	call dvbbs_error()
else%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center>

<tr> 
<td class=tablebody1 valign=middle height=50 align=center>
<iframe src="http://www.arkbook.com/icon/frame.html" width="97%" height="495" frameborder="no" border="0" marginwidth="0" marginheight="0" scrolling="no"></iframe>
</td>
</tr>
</table>
<br>
<%nam=checkstr(trim(request("nam1")))
if nam<>"" then
	call hy(nam)
end if
end if
call activeonline()
call footer()
%>