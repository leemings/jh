<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'监禁ip
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【监禁ip】</font><font color=" & saycolor & ">"+jianjin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'监禁ip
function jianjin(fn1,to1)
if aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：10级站长才能操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
if to1=Application("aqjh_user") then
	call boot(aqjh_name,"监禁ip，操作者："&aqjh_name&","&fn1)
	conn.execute "update 用户 set 状态='监禁',登录=now()+1,事件原因='监禁站长ip你找死！' where 姓名='" & aqjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
rs.open "select 姓名,grade,lastip from 用户 where 姓名='" & to1 &"'",conn,2,2
if aqjh_grade<=rs("grade") and aqjh_name<>Application("aqjh_user")  then
	Response.Write "<script language=JavaScript>{alert('提示：你不是主站长，不可以操作其它站长！');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
tolastip=rs("lastip")
conn.execute "update 用户 set 状态='监禁',登录=now()+1,事件原因='"&aqjh_name&" 监禁ip:"&to1&"不受江湖欢迎\n你们最后上线ip相同！' where grade<>10 and lastip='" & tolastip &"'"
jianjin="<font color=red><bgsound src=wav/daipu.wav loop=1>因为%%在江湖上不受欢迎，站长##封了与他相同登录ip的所有人……</font>"
call boot(to1,"监禁，操作者："&aqjh_name&","&fn1)
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','监禁',now(),'与" & to1 & "相同ip所有人监禁!')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>