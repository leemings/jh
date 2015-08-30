<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'封锁ip♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【封锁ip】</font><font color=" & saycolor & ">"+locklog(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'封锁ip
function locklog(fn1,to1)
if sjjh_grade<10 or sjjh_name<>"回首当年" then
	Response.Write "<script language=JavaScript>{alert('提示：封锁IP只有正站长才可以操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
if to1=Application("sjjh_user") then
	call boot(sjjh_name,"封锁ip，操作者："&sjjh_name&","&fn1)
	conn.execute "update 用户 set 状态='监禁',登录=now()+1,事件原因='封锁站长ip你找死！' where 姓名='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
rs.open "select 姓名,grade,lastip from 用户 where 姓名='" & to1 &"'",conn,3,3
if sjjh_grade<=rs("grade") and sjjh_name<>Application("sjjh_user")  then
	Response.Write "<script language=JavaScript>{alert('提示：你不是主站长，不可以操作其它站长！');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
tolastip=rs("lastip")
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
conn.Execute "INSERT INTO i (a,b,c) VALUES ('"& tolastip & "','"& sj & "','"&sjjh_name&"')"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','封ip30分钟："&lockip&"')"
locklog="封锁IP：(<font color=blue>" & tolastip & "</font>)被管理员封锁!"
call boot(to1,"封锁ip，操作者："&sjjh_name&","&fn1)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>