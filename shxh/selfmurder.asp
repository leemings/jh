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
rst.Open "select ��ż from �û� where �ʺ�='"&account&"' and ����='"&username&"' and ����='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=019"
mate=rst("��ż")
rst.Close
set rst=nothing
if mate<>"��" then conn.Execute "update �û� set ��ż='��' where ����='"&mate&"'"
conn.Execute "delete * from �û� where ����='"&username&"'"
conn.Execute "delete * from �ż� where ������='"&username&"'"
conn.Execute "delete * from ��Ʒ where ������='"&username&"'"
conn.Execute "delete * from ý�� where ������='"&username&"'"
conn.Execute "delete * from ���� where ����='"&username&"'"
conn.Execute "delete * from bbs where ����='"&username&"'"
session.Abandon 
conn.close
set conn=nothing
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='style.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><p align=center><font color=ff0000 size=5>�� �� ��</font><br>��ɱ��ɣ�<br><input type=button onclick=javascript:top.window.close(); value=' �� �� '></p></body>"
%>