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
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 配偶 from 用户 where 帐号='"&account&"' and 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=019"
mate=rst("配偶")
rst.Close
set rst=nothing
if mate<>"无" then conn.Execute "update 用户 set 配偶='无' where 姓名='"&mate&"'"
conn.Execute "delete * from 用户 where 姓名='"&username&"'"
conn.Execute "delete * from 信件 where 收信人='"&username&"'"
conn.Execute "delete * from 物品 where 所有者='"&username&"'"
conn.Execute "delete * from 媒婆 where 申请人='"&username&"'"
conn.Execute "delete * from 攻击 where 姓名='"&username&"'"
conn.Execute "delete * from bbs where 作者='"&username&"'"
session.Abandon 
conn.close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='style.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><p align=center><font color=ff0000 size=5>断 魂 涯</font><br>自杀完成！<br><input type=button onclick=javascript:top.window.close(); value=' 关 闭 '></p></body>"
%>