<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../config.asp"-->
<%'抓小偷
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<25 then
	Response.Write "<script language=JavaScript>{alert('提示：你的等级不够25级,也敢抓小偷！小心小偷把你也给偷了!');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【抓小偷】</font><font color=" & saycolor & ">"+zuaxiaotou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function zuaxiaotou(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
fromdj=rs("等级")
if to1="大家" or to1=aqjh_name or to1=Application("aqjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('抓小偷对象有错，请看仔细了！！');}</script>"
	Response.End
exit function
end if
if rs("职业")="小偷" then
zuaxiaotou="##:有没搞错呀，你就是小偷，还抓什么小偷呀！"
exit function
end if
rs.close
rs.open "SELECT 等级,门派,职业 FROM 用户 WHERE 姓名='"&to1&"'",conn
todj=rs("等级")
if rs("门派")="官府" then
   Response.Write "<script language=JavaScript>{alert('你有没有搞错呀，官府哪会有小偷!');}</script>"
   Response.End
end if
if rs("职业")<>"小偷" then
zuaxiaotou="##:你发什么神经啊！%%可没偷过东西，你抓他做什么？"
exit function
end if
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update 用户 set 职业='采冰' where 姓名='"&to1&"'"
	conn.execute "update 用户 set 道德=道德+200  where 姓名='"&aqjh_name&"'"
	zuaxiaotou="小偷[%%]在[##]的教导下,决定投案自首,改过自新不再做小偷![##]道德提升200点)"
Case 2
	if todj>fromdj then
		conn.execute "update 用户 set 魅力=魅力-100  where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 职业='采冰' where 姓名='"&to1&"'"
		zuaxiaotou="小偷[%%]武功高强,[##]根本不是小偷的对手,终于让小偷[%%]给跑了，[##]魅力下降100点)"
	else
		zuaxiaotou="[##]正想逮住小偷[%%],仔细一瞧,原来是自家亲戚,于是把小偷[%%]放了!"
	end if 
Case 3
	conn.execute "update 用户 set 魅力=魅力+100,道德=道德+100,银两=银两+10000,内力=内力-200 where 姓名='"&aqjh_name&"'"
	conn.execute "update 用户 set 职业='采冰',状态='死' where 姓名='"&to1&"'"
	call boot(to1,"抓小偷，操作者："&aqjh_name&","&fn1)
	zuaxiaotou="经过一番苦战,大好人[##]终于将小偷[%%]逮捕归案!(获得奖金10000两,道德提升100点,魅力提升100点,内力下降200点)"
Case 4
	conn.execute "update 用户 set 道德=道德-100,银两=银两+10000 where 姓名='"&aqjh_name&"'"
	conn.execute "update 用户 set 职业='采冰',银两=银两-10000 where 姓名='"&to1&"'"
	zuaxiaotou="[##]正想抓[%%]，一看原来认识，于是[%%]将偷来的银两分给[##]10000两，[##]也有些心动，反正偷的不是我的东西，就揣在兜里，扬长而去..."
case else
	zuaxiaotou="[##]正想逮住小偷{%%},仔细一瞧,原来是自家亲戚,于是把小偷{%%}放了!"
End Select
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>