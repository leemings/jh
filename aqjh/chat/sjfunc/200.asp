<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'嫁衣神功
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_jy then
	Response.Write "<script language=JavaScript>{alert('提示：嫁衣神功最少需要["&jhdj_jy&"]级,才可以操作！');}</script>"
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
says="<font color=green>【嫁衣神功】<font color=" & saycolor & ">"+cuan2(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'嫁衣神功
function cuan2(fn1,to1)
fn1=abs(fn1)
if fn1<10000 then
	Response.Write "<script language=JavaScript>{alert('提示：传送武功一次最少10000的，别太小气！');}</script>"
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
if fn1>50000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：传武功不要大于50000好不！');}</script>"
	Response.End
end if
if fn1<10000 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：传武功不要小于10000好！');}</script>"
	Response.End
end if
if rs("武功")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的武功不足！');}</script>"
	Response.End
end if
if rs("体力")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的体力不够100000');}</script>"
	Response.End
end if
if rs("内力")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的内力不够100000！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府的，你想做什么啊！');}</script>"
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
conn.execute "update 用户 set 知质=知质-100 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 体力=体力-8000 where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 内力=内力-15000 where 姓名='" &aqjh_name &"'"
conn.execute "update 用户 set 武功=武功+" & fn1*0.90 & " where 姓名='" & to1 &"'"
conn.execute "update 用户 set 武功加=武功加+20 where 姓名='" & to1 &"'"
cuan2= "##修炼的<bgsound src=wav/dz.wav loop=1>江湖第一神功，嫁衣神功，想不到尽是要传给别人才可以的<bgsound src=wav/dz.wav loop=1>，辛辛苦苦裁布做衣，竟是为她人做嫁妆，唉，##发功，终于把" & fn1 & "的武功传给了%%了，中途流失10%武功，知质下降100，%%万分感谢！##体力<font color=red>-8000</font>点,内力<font color=red>-15000</font>!%%武功上限提升<font color=red>+20</font>点!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>