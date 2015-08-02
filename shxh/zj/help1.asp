<%
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
if Application("Ba_jxqy_systemname")="" then Response.Redirect "error.asp?id=001"
if username="" or password="" then Response.Redirect "error.asp?id=002"
for i=1 to len(username)
	usernamechr=mid(username,i,1)
	if asc(usernamechr)>0 then Response.Redirect "error.asp?id=003"
next
if len(password)<6 then Response.Redirect "error.asp?id=004"
for i=1 to len(password)
	passwordchr=asc(mid(password,i,1))
	if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "error.asp?id=005"
next
lastip=Request.ServerVariables("REMOTE_ADDR")
lastlogintime=now()
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 门派,身份,性别,等级,道德,状态,最后登录时间,会员,会员时间 from 用户 where 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=010"
mycorp=rst("门派")
myidentity=rst("身份")
mysex=rst("性别")
mygrade=rst("等级")
moral=rst("道德")\255
predicament=rst("状态")
if rst("会员时间")>=date() and rst("会员")=true then
	fellow=true
else
	fellow=false
end if		
allowlogintime=rst("最后登录时间")
if predicament="死亡" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=018&allowtime="&allowlogintime
if predicament="逮捕" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="入狱" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="眠" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=036"
if moral>255 then moral=255
if moral<-255 then moral=-255
if moral>0 then
	mycolor=rgb(0,0,moral)
else
	mycolor=rgb(abs(moral),0,0)
end if
mycolor=cstr(hex(mycolor))
colorlen=len(mycolor)
select case colorlen
	case 1 
		mycolor="00000"&mycolor
	case 2
		mycolor="0000"&mycolor
end select			
rst.Close
set rst=nothing
conn.Execute "update 用户 set 最后登录IP='"&lastip&"',最后登录时间="&lastlogintimetype&",状态='正常',会员="&fellow&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
if session("Ba_jxqy_username")<>"" then Response.Redirect "myhome.asp"
Application.Lock
application("Ba_jxqy_allonline")=application("Ba_jxqy_allonline")&username&";"
session("Ba_jxqy_username")=username
session("Ba_jxqy_usersex")=mysex
session("Ba_jxqy_usercorp")=mycorp
Session("Ba_jxqy_useridentity")=myidentity
session("Ba_jxqy_usergrade")=mygrade
session("Ba_jxqy_userip")=lastip
session("Ba_jxqy_usermoral")=mycolor
session("Ba_jxqy_userchatroomsn")=1
session.Abandon
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
<LINK href="style1.css" rel=stylesheet>
</head>
<body bgcolor="<%=bgcolor%>" background="<%=bgimage%>"><center>
</body>  
<p align="center">已经成功处理非法掉线问题！</font></p>
