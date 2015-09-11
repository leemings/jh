<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'亲密接触
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_qr then
	Response.Write "<script language=JavaScript>{alert('提示：亲密接触需要["&jhdj_qr&"]级才可以操作！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：不会吧？想亲江湖所有的美女，不累死你才怪！');}</script>"
	Response.End
end if
if application("aqjh_tw")<>"" then
	wddata=split(application("aqjh_tw"),"|")
	if ubound(wddata)=2 then
		sj=wddata(2)
	end if
	erase wddata
	nowsj=DateDiff("s",sj,now())
	if nowsj<10 then
		jgsj=10-nowsj
		Response.Write "<script language=JavaScript>{alert('上一个亲密接触还未结束，请再等"& jgsj &"秒！');}</script>"
		Response.End
	end if
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
says="<font color=green>【亲密接触】<font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'求签
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,体力,内力,道德,魅力,性别 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<500 or rs("魅力")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不够500或魅力不够500，不可以色人！');}</script>"
	Response.End
end if
if rs("体力")<800 or rs("内力")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：亲密接触要耗废800体力和500内力，你太虚了吧！！');</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
'if sj<100 then
'	s=100-sj
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('亲密接触要100秒一次，请等"& s &"秒,可别累着！');</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：亲密接触只能找异性朋友!');}</script>"
	Response.End
end if
if rs("道德")<1 or rs("魅力")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方道德不够100或魅力不够100，色他（她）也太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<800 or rs("内力")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：亲密接触要耗废800体力和500内力，"& to1 &"太虚了！！');</script>"
	response.end
end if
conn.execute "update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_tw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
randomize()
regjm=int(rnd*3144998)
aa="<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value="
cc="onclick=javascript:;jgj"&regjm&".disabled=1;qjd"&regjm&".disabled=1;jhdt"&regjm&".disabled=1;jj"&regjm&".disabled=1;window.open('wd1.asp?zl="
tw="[##]向{%%}<img src='img/xu1.gif'>提出了亲密接触的请求：宝贝！来哟。我有许可证在手。再说我这么优秀！让我亲一亲值得啊！"& fn1 & aa &"'进房间' "& cc &"进房间','d') name=jgj"&regjm&">"& aa &"'去酒店' "& cc &"去酒店','d') name=qjd"&regjm&">"& aa &"'在江湖大厅' "& cc &"在江湖大厅','d') name=jhdt"&regjm&">"& aa &"'拒绝' "& cc &"拒绝','d') name=jj"&regjm&">"
end function
%>