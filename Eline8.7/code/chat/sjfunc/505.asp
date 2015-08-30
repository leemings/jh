<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'倒夜香♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_frag then
	Response.Write "<script language=JavaScript>{alert('提示：倒夜香需要["&jhdj_frag&"]级才可以操作！');}</script>"
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
call dianzan(towho)
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
says="<font color=green>【倒夜香】<font color=" & saycolor & ">"+frag(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'倒夜香
function frag(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 职业,内力,银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if trim(towho)="" or towho="大家" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
if rs("职业")<>"倒夜香" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的职业不是倒夜香，不能操作！');}</script>"
	Response.End
end if

if rs("银两")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没有2万块钱，不能倒夜香！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你内力不够1000不能倒夜香！');}</script>"
	Response.End
end if
rs.close
rs.open "select 银两,魅力,体力,内力,武功 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs("银两")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方没有2万块钱，不能帮他倒夜香！');}</script>"
	Response.End
end if
if rs("内力")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方内力低于1000，不能帮他倒夜香！');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
frag="倒夜香师<img src='pic/dz621.gif' width='40' height='55'>[##]替{%%}倒粪库：倒一次粪库收费2万银两内力1000喔<br><input style='font-size: 9pt; background-color: #99ccff; border-style: ridge' type=button value='立刻解决粪库' onClick=javascript:;tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('frag.asp?fromname=" & sjjh_name &"&toname="&to1&"','d') name=tongyi"&regjm&">" 


 
end function
%>
