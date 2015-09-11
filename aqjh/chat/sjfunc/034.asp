<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'吃饭
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_qk then
	Response.Write "<script language=JavaScript>{alert('提示：请客需要["&jhdj_qk&"]级才可以操作！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：不可以请所有人或自己呀！');}</script>"
	Response.End
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
says="<font color=blue>【掏钱请客】<font color=green>"+qw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'亲吻
function qw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,道德,魅力,性别 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<10 or rs("魅力")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德不够10或魅力不够50，不可以请对方！');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作时间"),now())
'if sj<20 then
'	s=20-sj
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('请人家要20秒一次，请等"& s &"秒再请,可别累着！');</script>"
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
if rs("体力")<800 or rs("内力")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：请客要耗废800体力和500内力，"& to1 &"太虚了！！');</script>"
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
cc="onclick=javascript:;ai"&regjm&".disabled=1;jj"&regjm&".disabled=1;window.open('hyqk.asp?zl="
qw="[##]捧着一大把钱向{%%}说："& fn1 & aa &"'吃饭' "& cc &"吃饭','d') name=ai"&regjm&">"& aa &"'冰淇淋' "& cc &"冰淇淋','d') name=ai"&regjm&">"& aa &"'点心' "& cc &"点心','d') name=ai"&regjm&">"& aa &"'拒绝' "& cc &"拒绝','d') name=jj"&regjm&">这些你任挑一种吧，我请你~！"
end function
%>