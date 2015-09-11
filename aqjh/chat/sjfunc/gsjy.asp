<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'观赏家园
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<40 then
	Response.Write "<script language=JavaScript>{alert('提示：邀请观赏家园需要40级才可以操作！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：请所有人来观赏家园，不累死你才怪！');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=1
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>〖邀请观赏家园〗<font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,银两,金币,体力,内力,道德,魅力,性别 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不够1000或魅力不够5000，不可以邀请观赏家园！');}</script>"
	Response.End
end if
'if rs("体力")<8 or rs("内力")<5 then
	'rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
	'Response.Write "<script Language=Javascript>alert('提示：体力不够1000或内力不够5000，不可以邀请观赏家园！');</script>"
	'response.end
'end if
if rs("银两")<200000 or rs("金币")<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：银两不够200000或金币不够2个，不可以邀请观赏家园！');</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<60 then
s=60-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('请人观赏家园要60秒一次，请等"& s &"秒,可别累着！');</script>"
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
'if rs("道德")<1000 or rs("魅力")<5000 then
	'rs.close
	'set rs=nothing
	'conn.close
	'set conn=nothing
	'Response.Write "<script language=JavaScript>{alert('提示：对方道德不够1000或魅力不够5000，邀请观赏家园也太没面子了！');}</script>"
	'Response.End
'end if
'if rs("体力")<8000 or rs("内力")<5000 then
	'rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
	'Response.Write "<script Language=Javascript>alert('提示：观赏家园要耗废8000体力和5000内力，"& to1 &"太虚了！！');</script>"
	'response.end
'end if

'conn.execute "update 用户 set 操作时间=now(),银两=银两-200000,金币=金币-2,体力=体力-8000,内力=内力-5000 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_sanghua")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
randomize()
regjm=int(rnd*3144998)
tw="<bgsound src=wav/ring.wav loop=1>[##]向{%%}提出了观赏家园的邀请:<b><a href=../jhshow/z_home_show.asp?name=" & aqjh_name & " target=_blank>来吧，朋友，点击进入我的虚拟家园！</a></b>"
end function
%>
