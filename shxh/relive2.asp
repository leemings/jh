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
rst.Open "select �ȼ�,״̬,����¼ʱ�� from �û� where �ʺ�='"&account&"' and ����='"&username&"' and ����='"&password&"'",conn
if rst.EOF or rst.BOF  then Response.Redirect "error.asp?id=019"
if rst("״̬")<>"����" then Response.Redirect  "error.asp?id=020"
on error resume next
conn.Execute "update �û� set ״̬='����',����=����*0.8,�ȼ�=�ȼ�-1,����=����*0.8,����=����*0.8,����=����*0.8,����=����*0.8 where ����='"&username&"' and ����='"&password&"' and �ʺ�='"&account&"' and ״̬='����'"
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
<p align="center">���Ѿ�����´α�����Ӵ</p>
<p align="right"><a href="relive2.asp#" onclick="javascript:window.close();">�ر�</a></p>
</body>