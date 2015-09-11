<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'请求亲吻
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_qr then
	Response.Write "<script language=JavaScript>{alert('提示：请求亲吻需要["&jhdj_qr&"]级才可以操作！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：不可以亲所有人或自己呀！');}</script>"
	Response.End
end if
if application("aqjh_qw")<>"" then
	wddata=split(application("aqjh_qw"),"|")
	if ubound(wddata)=2 then
		sj=wddata(2)
	end if
	erase wddata
	nowsj=DateDiff("s",sj,now())
	if nowsj<60 then
		jgsj=60-nowsj
		Response.Write "<script language=JavaScript>{alert('上一个刚亲还未结束，请再等"& jgsj &"秒！');}</script>"
		Response.End
	end if
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【抛个飞吻】<font color=" & saycolor & ">"+qw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'亲吻
function qw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,体力,内力,道德,魅力,性别 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不够1000或魅力不够5000，不可以亲吻对方！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：亲吻要耗废8000体力和5000内力，你太虚了吧！！');</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
'if sj<20 then
'	s=20-sj
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('亲人家要20秒一次，请等"& s &"秒再亲,可别累着！');</script>"
'	response.end
'end if
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
	Response.Write "<script language=JavaScript>{alert('提示：亲吻只能找异性朋友！难道你想同呀~晕哦');}</script>"
	Response.End
end if
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方道德不够1000或魅力不够5000，亲也太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：亲吻要耗废8000体力和5000内力，"& to1 &"太虚了！！');</script>"
	response.end
end if
conn.execute "update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_qw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
randomize()
regjm=int(rnd*3144998)
aa="<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value="
cc="onclick=javascript:;ai"&regjm&".disabled=1;jj"&regjm&".disabled=1;window.open('wenn.asp?zl="
qw="[##]对着{%%}飞了一个<img src='../chat/pic/dp76.gif'>"& fn1 & aa &"'飞吻' "& cc &"飞吻','d') name=ai"&regjm&">"& aa &"'亲亲' "& cc &"亲亲','d') name=ai"&regjm&">"& aa &"'狂吻' "& cc &"狂吻','d') name=ai"&regjm&">"& aa &"'拒绝' "& cc &"拒绝','d') name=jj"&regjm&">"
end function
%>