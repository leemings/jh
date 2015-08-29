<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'求签♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_chi then
	Response.Write "<script language=JavaScript>{alert('提示：求签需要["&jhdj_chi&"]级才可以操作！');}</script>"
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
says="<font color=green>【求签】<font color=" & saycolor & ">"+chice(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'求签
function chice(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 银两,职业,操作时间,体力,内力,道德,魅力 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德与魅力不够300，人家不会要跟你求签啦！');}</script>"
	Response.End
end if
if rs("银两")<200000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('给人算命的时候自个也多带点钱吧，当心算出下下签后让人家扁你！');}</script>"
	Response.End
end if
if rs("职业")<>"算命师" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的职业不是算命师，所以您不能在这里为他人进行算命！！');</script>"
	response.end
end if
if rs("体力")<800 or rs("内力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：算命要耗废800体力和1000内力，你太虚了吧！！');</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('算命求签60秒一次，请等"& s &"秒,可别累着！');</script>"
	response.end
end if
rs.close
if to1<>"大家" then
	rs.open "select 银两 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
	if rs("银两")<200000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：对方身上连20万块都没有，还给他算什么命呀。');}</script>"
		Response.End
	end if
	rs.close
end if
conn.execute "update 用户 set 操作时间=now() where 姓名='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application("sjjh_smxs")=sjjh_name&"|"&to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
if to1<>"大家" then
	chice="算命大师<img src='../picc/05.gif' width='40' height='55'>[##]替{%%}求神问卜：看看你今天的运气如何，求一支签吧，价格公道，收费合理，一次5万两，不准不要钱。<img src='../picc/chice.gif' width='60' height='50'><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='运程' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=运程','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='姻缘' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=姻缘','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='事业' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=事业','d') name=shiye"&regjm&">"
else
	chice="算命大师<img src='../picc/05.gif' width='40' height='55'>[##]在聊天室中喊到：谁要求签呀？价格公道，收费合理，一次5万两，不准不要钱。<img src='../picc/chice.gif' width='60' height='50'><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='运程' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=运程','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='姻缘' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=姻缘','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='事业' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=事业','d') name=shiye"&regjm&">"	
end if
end function
%>