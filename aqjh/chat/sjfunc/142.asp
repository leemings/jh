<!--#include file="sjfunc.asp"-->
<%'拧嘴
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
	Response.Write "<script language=JavaScript>{alert('提示：不会吧？你想对谁使用啊？');}</script>"
	Response.End
end if
act=1
towhoway=0
says="<font color=green>【家法惩治】<font color=" & saycolor & ">"+tw(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tw(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别,配偶,道德 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的道德太败坏了，不怕被人骂吗？');}</script>"
	Response.End
end if
mysex=rs("性别")
mywf=rs("配偶")
if mysex="男" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：这种功能是对你使用的啊！');}</script>"
	Response.End
end if
if mywf="无" then
 tw="##:你还没有结婚呢，不能使用家法惩治哦^_^"
 exit function
end if
if mywf<>to1 then
 tw="##:你不能对%%使用家法惩治，%%可不是你老公哦^_^"
 exit function
end if
rs.close
rs.open "select 体力,魅力 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方不是江湖中人，不可以对他使用此功能！');}</script>"
	Response.End
end if
if rs("魅力")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你老公的魅力已经很低了就放过这一次吧！');}</script>"
	Response.End
end if
if rs("体力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你老公的体力已经很弱了就放过这一次吧！');</script>"
	response.end
end if
conn.execute "update 用户 set 道德=道德-30,操作时间=now() where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 魅力=魅力-50,体力=体力-40 where 姓名='"&to1&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
tw="夫人##气哼哼地对老公%%使用了家法<img src=pic/zm1.gif>，看你还敢出去胡说八道！打得老公眼冒金星，满地找牙，体力下降40，魅力损失50，由于夫人温柔不再，道德下降30"
end function
%>