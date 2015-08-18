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
if not(sex="男" or sex="女") then Response.Redirect "error.asp?id=008"
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
	rst.Open "select 银两,注册IP,最后登录IP from 用户 where 姓名='"&recommender&"'",conn
	if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=074"
	recommendMoney = rst("银两")
	recommendIP = rst("注册IP")
	recommendLoginIP = rst("最后登录IP")
end if

regtime=now()
regip=Request.ServerVariables("REMOTE_ADDR")
on error resume next
conn.BeginTrans
conn.Execute "insert into 用户(帐号,密码,电子邮箱,contact,recommender,签名档,注册IP,注册时间,最后登录ip,最后登录时间,最后领钱日期,状态,姓名,性别,门派,身份,配偶,精力,等级,银两,积分,体力,内力,攻击,防御,资质,道德,特技,存款,结算日期,会员,会员时间,protect) values('"&account&"','"&password&"','"&e_mail&"','"&contact&"','"&recommender&"','"&sign&"','"&regip&"','"&regtime&"','"&regip&"','"&regtime&"','"&regtime&"','正常','"&username&"','"&sex&"','无','无','无',0,1,100,0,100,100,10,10,0,100,';',0,'"&regtime&"',false,'"&regtime&"','"&regtime&"')"
if conn.Errors.Count=0 then
	conn.CommitTrans
	If recommender<>"" then
		nowtimetype="#"&month(regtime)&"/"&day(regtime)&"/"&year(regtime)&" "&hour(regtime)&":"&minute(regtime)&":"&second(regtime)&"#"
		if recommendIP = regip or recommendLoginIP = regip then
			conn.execute "insert into 信件(收信人,标题,内容,写信人,写信时间,观看) values('"&recommender&"','拉人有奖','恭喜你拉来了〖"&username&"〗，但是别用小号刷奖励啦','系统',"&nowtimetype&",False)"
		else
			recommendMoneyNew = recommendMoney + 1000000
			conn.Execute "update 用户 set 银两='"&recommendMoneyNew&"' where 姓名='"&recommender&"'"
			conn.execute "insert into 信件(收信人,标题,内容,写信人,写信时间,观看) values('"&recommender&"','拉人有奖','恭喜你拉来了〖"&username&"〗，特意奖励你100000两银子，拉来的朋友，每升级到3，6，9，12级，你都会获得奖励, 请再接再励，江湖有你更精彩。','系统',"&nowtimetype&",False)"
		end If
	End if		
else	
	conn.RollbackTrans
	if conn.Errors(0).Number =-2147467259 then Response.Redirect "error.asp?id=009"
	Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
end if

conn.Close
set conn=nothing
response.write "<head><link rel=stylesheet href=style.css></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' topmargin=100><div align=center><font color=0000FF size=6>注册成功，请牢记帐号和密码</font><p><input type='button' value='关闭' onclick='javascript:top.window.close();' id=button1 name=button1></div>"
%>