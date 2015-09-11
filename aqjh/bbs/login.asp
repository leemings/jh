<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
top

select case Request("menu")

case "password"


%>



<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 验证密码</td>
</tr>
</table>
<br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>

<form action="login.asp" method="post">
<input type="hidden" value="passwordok" name="menu">
<input type="hidden" value="<%=Request("url")%>" name="url">
<tr>
<td width="328" height="25" align="center" class=a1>
登录论坛
</td></tr><tr>
<td height="19" width="328" valign="top" align="center" class=a3>
通行密码: <input type="password" size="15" name="password"><br>

<input type="submit" value=" 登录 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit">
</td></tr> </form></table>

<br>
<center>
<a href=javascript:history.back()>BACK </a><br>


<%
htmlend
case "passwordok"
Response.Cookies("password")=Request("password")
response.redirect Request("url")

case "add"


url=Request("url")
username=HTMLEncode(Trim(Request("username")))
userpass=md5(Trim(Request("userpass")))
if username=empty then error("<li>用户名没有输入")

sql="select * from [user] where username='"&username&"'"
Set Rs1=Conn.Execute(SQL)
if rs1.eof then error("<li>此用户名还未注册")

if Len(rs1("userpass"))<16 then
if Request("userpass")<>rs1("userpass") then error("<li>您输入的密码错误")
conn.execute("update [user] set userpass='"&userpass&"' where username='"&username&"'")

elseif Len(rs1("userpass"))=16 then
mdfive=16
if md5(Request("userpass"))<>rs1("userpass") then error("<li>您输入的密码错误")
conn.execute("update [user] set userpass='"&userpass&"' where username='"&username&"'")

else
if userpass<>rs1("userpass") then error("<li>您输入的密码错误")
end if


if rs1("membercode")=0 then error("<li>您的帐号尚未激活")
Response.Cookies("username")=username
Response.Cookies("userpass")=userpass
Response.Cookies("eremite")="0"

if Request("eremite")="1" then Response.Cookies("eremite")="1"

if Request("xuansave")=1 then
Response.Cookies("eremite").Expires=date+365
Response.Cookies("username").Expires=date+365
Response.Cookies("userpass").Expires=date+365
end if

if url<>empty and instr(url,"login.asp")=0 and instr(url,"left.asp")=0 and instr(url,"error.asp")=0 then
response.redirect url
else
response.write "<SCRIPT>location='Default.asp';</SCRIPT>"
end if

case "out"
conn.execute("delete from [online] where sessionid='"&session.sessionid&"'")
Response.Cookies("username")=""
Response.Cookies("userpass")=""
succtitle="已经成功退出"
message=message&"<li><a href=Default.asp>社区首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=Default.asp>")

end select
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> →  登录社区</td>
</tr>
</table><br>
<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>

<form action="login.asp" method="post">
<input type="hidden" value="add" name="menu">
<input type="hidden" value="<%=Request.ServerVariables("HTTP_REFERER")%>" name="url">
<tr>
<td width="328" height="25" align="center" class=a1>
登录社区
</td></tr><tr>
<td height="19" width="328" valign="top" align="center" class=a3>
用户名称:
<input size="15" name="username" value="<%=Request.Cookies("username")%>" readonly>&nbsp; <a href="register.asp">没有注册?</a><br>
用户密码: <input type="password" size="15" value name="userpass">&nbsp;
<a href="RecoverPasswd.asp">忘记密码?</a><br>

<input type="checkbox" value="1" name="xuansave" id="xuansave"><label for="xuansave">记住密码</label>

<input type="checkbox" value="1" name="eremite" id="eremite"><label for="eremite">隐身登录</label><br>
<input type="submit" value=" 登录 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit">
</td></tr> </form></table>

<br>
<center>
<a href=javascript:history.back()>BACK </a><br>


<%htmlend%>