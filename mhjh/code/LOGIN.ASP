
<!--#include file="pass.asp"--><%
reg=Request.form("reg")
reg1=Request.form("reg1")
if reg<>reg1 then
%>
<script language=vbscript>
MsgBox "你输入的认证码不正确，应该输入:<%=reg1%>"
location.href ="javascript:history.back()"
</script>
<%
else
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
lip=Application("yx8_mhjh_lockip")
if instr(username,"'")<>0  then Response.Redirect "error.asp?id=524"
if Application("yx8_mhjh_systemname")="" then Response.Redirect "error.asp?id=001"
if username="" or password="" then Response.Redirect "error.asp?id=002"
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0  or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"") then Response.Redirect "error.asp?id=056"
next
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
if instr(lip,";"&lastip&";")<>0 then Response.Redirect "error.asp?id=247"
password=md5(password)
lastlogintime=now()
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
	n=Year(date())
	y=Month(date())
	r=Day(date())
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	if len(y)=1 then y="0" & y
	if len(r)=1 then r="0" & r
	if len(s)=1 then s="0" & s
	if len(f)=1 then f="0" & f
	if len(m)=1 then m="0" & m
	sj=n & "-" & y & "-" & r 
%>
<!--#include file="data.asp"-->
<%
	sql="select * from 贷款 where  贷款人='"&username&"' and 还贷记录=false and DateDiff('d',贷款日期,#" & sj & "#)>7"
	Set Rs=connt.Execute(sql)
	if not(rs.BOF or rs.EOF) then 
	connt.Execute ("update 贷款 set 还贷记录=true where 贷款人='"&username&"'")
	connt.Close
	set connt=nothing
	conn.Execute "update 用户 set 体力=体力-100000,内力=内力-100000,状态='鬼魂' where 姓名='"&username&"'"
	conn.Close
	set conn=nothing
	Response.Redirect "error.asp?id=243"
	Response.end
	end if
rst.Open "select 门派,等级,状态,形态,最后登录时间,会员,会员时间 from 用户 where 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=010"
mycorp=rst("门派")
mygrade=rst("等级")
predicament=rst("状态")
xingtai=rst("形态")
if rst("会员时间")>=date() and rst("会员")=true then
fellow=true
else
fellow=false
end if
allowlogintime=rst("最后登录时间")
rst.Close
set rst=nothing
'if datediff("s",allowlogintime,lastlogintime)<0 then
't=abs(datediff("s",allowlogintime,lastlogintime))
'Response.write "<script language='javascript'>alert ('提示：你刚掉线了，要再过"&t&"秒才能登陆');location.href = 'index.asp';</script>"
'Response.End 
'end if
if predicament="逮捕" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="入狱" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=022&allowtime="&allowlogintime
if predicament="眠" and datediff("s",allowlogintime,lastlogintime)<0 then Response.Redirect "error.asp?id=036"
if predicament="死亡" then Response.Redirect "error.asp?id=018"
if xingtai="鬼魂" then
conn.Execute "update 用户 set 体力=体力*0.95,内力=内力*0.95,最后登录IP='"&lastip&"',状态='正常',最后登录时间="&lastlogintimetype&" where 姓名='"&username&"'"
else
conn.Execute "update 用户 set 最后登录IP='"&lastip&"',最后登录时间="&lastlogintimetype&",状态='正常',会员="&fellow&" where 姓名='"&username&"'"
end if
'if session("yx8_mhjh_username")<>"" and session("yx8_mhjh_usercorp")<>"" then 
'session("yx8_mhjh_userchatroomsn")=1
'Response.Redirect "main.asp"
'end if
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
for i=1 to onlinelistubd step 6
   if onlinelist(i)=username then
   online=True
   end if
next
if instr(Application("yx8_mhjh_allonline"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename1"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename2"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename3"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename4"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename5"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename6"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename7"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename8"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename9"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename10"),";"&username&";")<>0 or instr(Application("yx8_mhjh_onlinename11"),";"&username&";")<>0 then 
Application.Lock
Application("yx8_mhjh_allonline")=replace(application("yx8_mhjh_allonline"),";"&username&";",";")
Application.UnLock
session("Ba_jxqy_username")=username
session("Ba_jxqy_usercorp")=mycorp
session("Ba_jxqy_usergrade")=mygrade
session("Ba_jxqy_userchatroomsn")=1
diaoxiantime=180
nowtime=now()
lastlogintime=dateadd("s",diaoxiantime,nowtime)
lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
conn.Execute "update 用户 set 最后登录时间="&lastlogintimetype&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
session.Abandon
Response.Redirect "error.asp?id=017"
Response.End 
end if
conn.Close
set conn=nothing
Application.Lock
application("yx8_mhjh_allonline")=application("yx8_mhjh_allonline")&username&";"
session("yx8_mhjh_username")=username
session("yx8_mhjh_usercorp")=mycorp
session("yx8_mhjh_usergrade")=mygrade
session("yx8_mhjh_userchatroomsn")=1
if predicament="鬼魂" and session("yx8_mhjh_username")<>"" and session("yx8_mhjh_usercorp")<>"" then
Response.write "<script language='javascript'>alert ('提示：由于你在江湖纷争中被杀死，你现在已经成了鬼魂，无法再打仗，也不会被人打了；唯一的拯救办法就是找个道德高深的和尚帮你超度成人，否则，每次登陆体力都会减少5%，有够惨的了！');location.href = 'main.asp';</script>"
Response.End 
elseif predicament="正常" and session("yx8_mhjh_username")<>"" and session("yx8_mhjh_usercorp")<>"" then
Response.Redirect "main.asp"
else
Response.Redirect "error.asp?id=017"
end if
end if
%>


