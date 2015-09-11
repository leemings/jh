<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'暗然销魂
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_dt then
	Response.Write "<script language=JavaScript>{alert('提示：暗然销魂最少需要["&jhdj_dt&"]级,才可以操作！');}</script>"
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
'对暂离开、点哑穴判断
call dianzan(towho)
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
says="<font color=green>【暗然销魂】<font color=" & saycolor & ">"+cuan2(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'暗然销魂
function cuan2(fn1,to1)
fn1=abs(fn1)
if fn1<200 then
	Response.Write "<script language=JavaScript>{alert('提示：暗然销魂武功一次最少200的，不然力气太小！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
cy=DateDiff("n",rs("操作时间"),now())
if cy<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-cy
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if fn1>30000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：暗然销魂攻击不要大于30000好不！不然人家必死');}</script>"
	Response.End
end if
if fn1<200 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：暗然销魂攻击不要小于200好！不然人家会反击你的');}</script>"
	Response.End
end if
if rs("武功")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的武功不足！无法使用暗然销魂攻击');}</script>"
	Response.End
end if
if rs("体力")<8000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的体力不够8000！无法使用暗然销魂攻击');}</script>"
	Response.End
end if
if rs("内力")<8000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的内力不够8000！无法使用暗然销魂攻击');}</script>"
	Response.End
end if
if rs("等级")<=30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是新手的，你想做什么啊！');}</script>"
	Response.End
end if
if rs("知质")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的知质不够1000！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 武功=武功-" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 知质=知质-1000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 体力=体力-8000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 内力=内力-8000 where 姓名='" &aqjh_name &"'"
conn.execute "update 用户 set 武功=武功-" & fn1*0.90 & " where 姓名='" & to1 &"'"
conn.execute "update 用户 set 武功加=武功加-10 where 姓名='" & to1 &"'"
cuan2= "##修炼的<bgsound src=wav/dz.wav loop=1>江湖情侣神功分手后自创的一种武功，暗然销魂，这招使出想像对方必定重伤，众大侠只见##脸面红红，##开始发功，终于用" & fn1 & "武功的对%%进行强攻，失去了自身的10%武功，知质下降1000，%%恼羞成怒，##体力<font color=red>-8000</font>点,内力<font color=red>-8000</font>!%%的武功下降不少，武功上限也降低了<font color=red>-10</font>点!这招真绝！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>