<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
for each element in Request.Form
	elevalue=Request.Form(element)
	if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") then Response.Redirect "error.asp?id=056"
next
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
repassword=trim(Request.Form("repassword"))
contact=trim(Request.Form("contact"))
recommender=trim(Request.Form("recommend"))
sex=Request.Form("sex")
e_mail=server.HTMLEncode(trim(Request.Form("e_mail")))
sign=server.HTMLEncode(trim(Request.Form("sign")))
if Application("Ba_jxqy_systemname")="" then Response.Redirect "error.asp?id=001"
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
if(len(e_mail)<6 or instr(e_mail,".")=0 or instr(e_mail,"@")=0) then Response.Redirect "error.asp?id=055"
if password<>repassword then Response.Redirect "error.asp?id=007"
for i=1 to len(password)
	passwordchr=asc(mid(password,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
if not(sex="��" or sex="Ů") then Response.Redirect "error.asp?id=008"
illegidimacyname=split(Application("Ba_jxqy_illegidimacyname"),";")
if IsArray(illegidimacyname) then
	for i=1 to ubound(illegidimacyname)-1
		if instr(username,illegidimacyname(i))<>0 then Response.Redirect "error.asp?id=057"
	next
end if

set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")

if recommender<>"" then
	set rst=server.CreateObject("adodb.recordset")
	rst.Open "select ����,ע��IP,����¼IP from �û� where ����='"&recommender&"'",conn
	if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=074"
	recommendMoney = rst("����")
	recommendIP = rst("ע��IP")
	recommendLoginIP = rst("����¼IP")
end if

regtime=now()
regip=Request.ServerVariables("REMOTE_ADDR")
on error resume next
conn.BeginTrans
conn.Execute "insert into �û�(�ʺ�,����,��������,contact,recommender,ǩ����,ע��IP,ע��ʱ��,����¼ip,����¼ʱ��,�����Ǯ����,״̬,����,�Ա�,����,���,��ż,����,�ȼ�,����,����,����,����,����,����,����,����,�ؼ�,���,��������,��Ա,��Աʱ��,protect) values('"&account&"','"&password&"','"&e_mail&"','"&contact&"','"&recommender&"','"&sign&"','"&regip&"','"&regtime&"','"&regip&"','"&regtime&"','"&regtime&"','����','"&username&"','"&sex&"','��','��','��',0,1,100,0,100,100,10,10,0,100,';',0,'"&regtime&"',false,'"&regtime&"','"&regtime&"')"
if conn.Errors.Count=0 then
	conn.CommitTrans
	If recommender<>"" then
		nowtimetype="#"&month(regtime)&"/"&day(regtime)&"/"&year(regtime)&" "&hour(regtime)&":"&minute(regtime)&":"&second(regtime)&"#"
		if recommendIP = regip or recommendLoginIP = regip then
			conn.execute "insert into �ż�(������,����,����,д����,д��ʱ��,�ۿ�) values('"&recommender&"','�����н�','��ϲ�������ˡ�"&username&"�������Ǳ���С��ˢ������','ϵͳ',"&nowtimetype&",False)"
		else
			recommendMoneyNew = recommendMoney + 1000000
			conn.Execute "update �û� set ����='"&recommendMoneyNew&"' where ����='"&recommender&"'"
			conn.execute "insert into �ż�(������,����,����,д����,д��ʱ��,�ۿ�) values('"&recommender&"','�����н�','��ϲ�������ˡ�"&username&"�������⽱����100000�����ӣ����������ѣ�ÿ������3��6��9��12�����㶼���ý���, ���ٽ�������������������ʡ�','ϵͳ',"&nowtimetype&",False)"
		end If
	End if		
else	
	conn.RollbackTrans
	if conn.Errors(0).Number =-2147467259 then Response.Redirect "error.asp?id=009"
	Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
end if

conn.Close
set conn=nothing
response.write "<head><link rel=stylesheet href=style.css></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' topmargin=100><div align=center><font color=0000FF size=6>ע��ɹ������μ��ʺź�����</font><p><input type='button' value='�ر�' onclick='javascript:top.window.close();' id=button1 name=button1></div>"
%>