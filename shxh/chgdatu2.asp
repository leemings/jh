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
rst.Open "select * from 用户 where 帐号='"&account&"' and 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF  then Response.Redirect "error.asp?id=019"
rst.Close
set rst=nothing
e_mail=server.HTMLEncode(trim(Request.Form("e_mail")))
sign=server.HTMLEncode(trim(Request.Form("sign")))
if instr(e_mail,".")=0 or instr(e_mail,"@")=0 or instr(e_mail,"'")<>0 or instr(e_mail,chr(34))<>0 or instr(e_mail,"\")<>0 then Response.Redirect "error.asp?id=055"
if instr(sign,"file:")<>0 or instr(sign,"script:")<>0 or instr(sign,"'")<>0 or instr(sign,chr(34))<>0 or instr(sign,"\")<>0 then Response.Redirect "error.asp?id=037"
conn.Execute "update 用户 set 电子邮箱='"&e_mail&"',签名档='"&sign&"' where 姓名='"&username&"'"
conn.Close
set conn=nothing
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>' oncontextmenu=self.event.returnValue=false topmargin="100">
<p align=center><font color=ff0000 size=5>个人资料更新完成！</font></p>
<p align=center><input type="button" value=" 关 闭 " onclick="javascript:top.window.close();" class=norsty></p>
</body>