<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'邀请跳舞♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_tw then
	Response.Write "<script language=JavaScript>{alert('提示：邀请舞伴需要["&jhdj_chi&"]级才可以操作！');}</script>"
	Response.End
end if
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
if towho=sjjh_name or towho=application("sjjh_automanname") then
	towho="大家"
else
	call dianzan(towho)
end if
if towho="大家" then
	Response.Write "<script language=JavaScript>{alert('提示：请所有人来作你的舞伴，不累死你才怪！');}</script>"
	Response.End
end if
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
says="<font color=green>【跳舞】<font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'求签
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间,体力,内力,道德,魅力,性别 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不够1000或魅力不够5000，不可以邀请舞伴！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：跳舞要耗废8000体力和5000内力，你太虚了吧！！');</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<100 then
s=100-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('请人跳舞要300秒一次，请等"& s &"秒,可别累着！');</script>"
response.end
end if
mysex=rs("性别")
rs.close
rs.open "select 体力,内力,道德,魅力,性别 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
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
	Response.Write "<script language=JavaScript>{alert('提示：跳舞只能找异性朋友！');}</script>"
	Response.End
end if
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方道德不够1000或魅力不够5000，邀请舞伴也太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：跳舞要耗废8000体力和5000内力，"& to1 &"太虚了！！');</script>"
	response.end
end if

conn.execute "update 用户 set 操作时间=now() where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_tw")=sjjh_name&"|"&to1
Application.UnLock
randomize()
regjm=int(rnd*3144998)
tw="<bgsound src=wav/ring.WAV loop=1>[##]向{%%}<img src='img/yu.gif' width='70' height='55'>提出了跳舞的请求，来哟，宝贝，门票才五万银两，便宜得很：<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='华尔兹' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=华尔兹','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='探戈' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=探戈','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='迪斯高' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=迪斯高','d') name=shiye"&regjm&">"
end function
%>