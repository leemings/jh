<%
for each element in Request.Form
	elevalue=Request.Form(element)
	if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") then Response.Redirect "error.asp?id=056"
next
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
rst.Open "select 等级,状态,最后登录时间 from 用户 where 帐号='"&account&"' and 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF  then Response.Redirect "error.asp?id=019"
if rst("状态")<>"死亡" then Response.Redirect  "error.asp?id=020"
on error resume next
conn.Execute "update 用户 set 状态='正常',积分=积分*0.8,等级=等级-1,体力=体力*0.8,内力=内力*0.8,攻击=攻击*0.8,防御=防御*0.8 where 姓名='"&username&"' and 密码='"&password&"' and 帐号='"&account&"' and 状态='死亡'"
if conn.Errors.Count=0 then
	conn.CommitTrans
else
	conn.RollbackTrans
	Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
end if
set rst=nothing	
conn.Close
set conn=nothing
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><link rel="stylesheet" href="style.css"></head>
<body bgcolor="<%=bgcolor%>" text="#000000" background="<%=bgimage%>">
<p align="center">你已经复活，下次别来了哟</p>
<p align="right"><a href="relive2.asp#" onclick="javascript:window.close();">关闭</a></p>
</body>