<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
    server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
    if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "你提交的路径有误，禁止从站点外部提交数据！"
        response.end
    end if
reg=Request.form("reg")
reg1=Request.form("reg1")
if reg<>reg1 then
%>
<script language=vbscript>
MsgBox "你输入的认证码不正确，应该输入:<%=reg1%>"
location.href = "javascript:history.back()"
</script>
<%
else
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,"?")<>0 or instr(elevalue,";") or instr(elevalue,"file:")<>0 or instr(elevalue,"script:")<>0 or instr(elevalue,"#")<>0 then Response.Redirect "../error.asp?id=056"
next
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
if account="" or username="" or reg="" then Response.Redirect "../error.asp?id=002"
if len(account)<6 or len(username)<2  then Response.Redirect "../error.asp?id=004"
if len(account)>15 or len(username)>7  then Response.Redirect "../error.asp?id=504"
if instr(account,"file:")<>0 or instr(account,"script:")<>0 or instr(account,"'")<>0 or instr(account,chr(34))<>0 or instr(account,"\")<>0 or InStr(account,",")<>0 or InStr(account,"<")<>0 or InStr(account,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(username,"file:")<>0 or instr(username,"script:")<>0 or instr(username,"'")<>0 or instr(username,chr(34))<>0 or instr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then Response.Redirect "../error.asp?id=524"
for i=1 to len(account)
accountchr=asc(mid(account,i,1))
if accountchr<48 or (accountchr>57 and accountchr<65) or (accountchr>90 and accountchr<97) then Response.Redirect "../error.asp?id=006"
next
for i=1 to len(username)
usernamechr=mid(username,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=003"
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 密码 from 用户 where 帐号='"&account&"' and 姓名='"&username&"'",conn
if rst.EOF or rst.BOF  then Response.Redirect "../error.asp?id=019"
mm=rst("密码")
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title><%=Application("yx8_mhjh_systemname")%></title><link rel="stylesheet" href="../style.css"></head>
<body  text="#000000" background='../chatroom/bg1.gif'>
<p align="center">你的密码是<%=mm%><br>要记住啊！再忘记就打你PP</p>
<p align="center"><a href="getpw2.asp#" onclick="javascript:window.close();">关闭</a></p>
</body>
<%end if%>