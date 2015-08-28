<%@ LANGUAGE=VBScript%><!--#include file="pass.asp"-->
<%Response.Expires=0
Response.Buffer=true
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
location.href ="javascript:history.back()"
</script>
<%
elseif reg1="" then
%>
<script language=vbscript>
MsgBox "认证码不能为空"
location.href ="javascript:history.back()"
</script>
<%
else
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
if Application("yx8_mhjh_systemname")="" then Response.Redirect "../error.asp?id=001"
if username="" or password="" then Response.Redirect "../error.asp?id=002"
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"") then Response.Redirect "../error.asp?id=056"
next
for i=1 to len(username)
	usernamechr=mid(username,i,1)
	if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=003"
next
if len(password)<6 then Response.Redirect "../error.asp?id=004"
for i=1 to len(password)
	passwordchr=asc(mid(password,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "../error.asp?id=005"
next
password=md5(password)
lastip=Request.ServerVariables("REMOTE_ADDR")
diaoxiantime=5
nowtime=now()
lastlogintime=dateadd("s",diaoxiantime,nowtime)
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 门派,性别,等级,状态,最后登录时间 from 用户 where 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=010"
mycorp=rst("门派")
mysex=rst("性别")
mygrade=rst("等级")
predicament=rst("状态")	
allowlogintime=rst("最后登录时间")
if predicament="逮捕" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "../error.asp?id=022&allowtime="&allowlogintime
if predicament="入狱" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "../error.asp?id=022&allowtime="&allowlogintime
if predicament="眠" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "../error.asp?id=036"
if predicament="死亡" then Response.Redirect "../error.asp?id=018"
rst.Close
set rst=nothing
conn.Execute "update 用户 set 最后登录IP='"&lastip&"',最后登录时间="&lastlogintimetype&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
Application.Lock
Application("yx8_mhjh_allonline")=replace(application("yx8_mhjh_allonline"),";"&username&";",";")
Application.UnLock
session("yx8_mhjh_username")=username
session("yx8_mhjh_usercorp")=mycorp
session("yx8_mhjh_usergrade")=mygrade
session("yx8_mhjh_userchatroomsn")=1
session.Abandon
%>
<head>
<title><%=Application("yx8_mhjh_systemname")%></title>
<LINK href="css.css" rel=stylesheet>
</head>
<body bgcolor="#800000" text="#000000"><center>
</center>  
<p align="center"><font color="#FFFFFF">已经成功处理非法掉线问题！请5秒钟后再登陆江湖</font></p>
</body>  
<%end if%>
































