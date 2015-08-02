<%
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
if account="" or username="" or password="" then Response.Redirect "error.asp?id=002"
if len(account)<6 or len(username)<2 or len(password)<6 then Response.Redirect "error.asp?id=004"
for i=1 to len(account)
	accountchr=asc(mid(account,i,1))
	if accountchr<48 or (accountchr>57 and accountchr<65) or (accountchr>90 and accountchr<97) then Response.Redirect "error.asp?id=006"
next	
for i=1 to len(username)
	usernamechr=mid(username,i,1)
	if asc(usernamechr)>0 then Response.Redirect "error.asp?id=003"
next
for i=1 to len(password)
	passwordchr=asc(mid(password,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select * from 用户 where 帐号='"&account&"' and 姓名='"&username&"' and 密码='"&password&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF  then Response.Redirect "error.asp?id=019"
msg="<tr><td><INPUT TYPE='hidden' NAME='account' VALUE='"&account&"'><INPUT TYPE='hidden' NAME='username' VALUE='"&username&"'><INPUT TYPE='hidden' NAME='password' VALUE='"&password&"'>电子邮箱</td><td><INPUT class=norsty maxLength=30 name=e_Mail size=30 value="&rst("电子邮箱")&"></td></tr><tr><td>签名档</td><td><INPUT class=norsty maxLength=100 name=sign size=40 value="&rst("签名档")&"></td></tr>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><link rel="stylesheet" href="style.css"></head>
<body bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>' oncontextmenu=self.event.returnValue=false>
<form action=chgdatu2.asp method=post>
<table border=3 width=90% align=center><%=msg%><tr align=center><td colspan=2><input type=submit value=" 更 改 "> <input type=button onclick=javascript:history.back(); value=' 返 回 '></td></tr></table>
</form>
</body>