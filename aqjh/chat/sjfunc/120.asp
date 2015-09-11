<!--#include file="sjfunc.asp"-->
<%'抢亲
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<100 then
	Response.Write "<script language=JavaScript>{alert('提示：抢亲需要100级才可以操作！');}</script>"
	Response.End
end if
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="大家"
else
	call dianzan(towho)
end if
if towho="大家" then
	Response.Write "<script language=JavaScript>{alert('提示：不会吧？抢大家？！');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","＆")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><b>【抢亲】</b><font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'抢亲
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
bl=rs("配偶")
if bl<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有没有搞错啊？都结过婚了，还抢个屁呀，嘿嘿！');}</script>"
	Response.End
end if
if rs("道德")<50000 or rs("魅力")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的道德或魅力少于5万……');}</script>"
	Response.End
end if
if rs("知质")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的知质少于500……');}</script>"
	Response.End
end if
mysex=rs("性别")
rs.close
rs.open "select * FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方不是江湖中人，不可以对他使用此功能！');}</script>"
	Response.End
end if
if rs("性别")=mysex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：晕倒！对方和你同一性别，难道你是……');}</script>"
	Response.End
end if
if rs("配偶")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方已经成家了，你就死了这条心吧……');}</script>"
	Response.End
end if
conn.execute "update 用户 set 配偶='"&aqjh_name&"' where 姓名='"&to1&"'"
conn.execute "update 用户 set 配偶='"&to1&"',allvalue=allvalue-int(allvalue/100),道德=道德-50000,魅力=魅力-50000,知质=知质-500 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
Application.Lock
Application("aqjh_tw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
tw="<font color=green>现在这是什么世道呀，[<font color=red><b>"&aqjh_name&"</b></font>]竟然不顾世俗眼光，将[<font color=red><b>"&to1&"</b></font>]强抢回家，哎！[<font color=red><b>"&aqjh_name&"</b></font>]积分一下子少了[<font color=red><b>百分之一</b></font>]，同时也失去了不少道德与魅力，值得吗？<font>"
end function
%>