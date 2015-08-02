<%
for each element in Request.Form
	elevalue=Request.Form(element)
	if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") then Response.Redirect "error.asp?id=056"
next
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
password1=trim(Request.Form("password1"))
password2=trim(Request.Form("password2"))
if account="" or username="" or password="" then Response.Redirect "error.asp?id=002"
if len(account)<6 or len(username)<2 or len(password)<6 or len(password1)<6 then Response.Redirect "error.asp?id=004"
if password1<>password2 then Response.Redirect "error.asp?id=007"
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
for i=1 to len(password1)
	passwordchr=asc(mid(password1,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
on error resume next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select id from 用户 where 姓名='"&username&"' and 密码='"&password&"' and 帐号='"&account&"'",conn
if rst.EOF or rst.BOF then
	Response.Redirect "error.asp?id=010"
else	
	conn.BeginTrans
	conn.Execute "update 用户 set 密码='"&password1&"' where 姓名='"&username&"' and 密码='"&password&"' and 帐号='"&account&"'"
	if conn.Errors.Count=0 then
		conn.CommitTrans
	else
		conn.RollbackTrans
		Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
	end if
end if	
conn.Close
set conn=nothing
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='style.css'></head><body topmargin='100' bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><p align=center><font color=ff0000 size=5>更改密码成功，为了您的帐号安全，您应该经常更改您的密码</font><br><input type=button onclick='javascript:top.window.close();' value=' 关 闭 ' ></p></body>"
%>