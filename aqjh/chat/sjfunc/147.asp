<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'盟主令
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_name<>Application("aqjh_mengzhu") then
	Response.Write "<script language=JavaScript>{alert('提示：你不是盟主，瞎捣什么乱啊？');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=mzl(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'盟主令
function mzl(fn1,to1)
saysok="<img src=pic/mengzhulin.gif><font color=ff0013>【盟主令】</font>"
'判断对方
if to1=aqjh_name or to1="大家" or to1=Application("aqjh_automanname") then
 	Response.Write "<script language=JavaScript>{alert('警告：操作对谁？');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='"&to1&"'",conn,1,3
to_grade=rs("grade")
to_tj=rs("通缉")
to_yin=rs("银两")
rs.close
if to_grade>5 then
 mzl=saysok&"##号令%%失败，因为%%是官府中人"
 exit function
end if
fn1=trim(fn1)
if fn1="帮助" or fn1="" then
	Response.Write "<script language=JavaScript>{alert('盟主令功能有:帮助／奖励／惩罚／斩首／踢人[不得乱用，否则删除id]');}</script>"
	Response.End
elseif fn1="奖励" then
 if to_tj=true then
  mzl=saysok&"武林盟主##对%%进行奖励失败，对方正在通缉中啊！" 
  exit function
 end if
 rs.open "select 操作时间 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,1,3
 sj=DateDiff("s",rs("操作时间"),now())
 if sj<55 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=55-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
 end if
 randomize()
 rnd1=int(rnd*400000)+100000
 conn.execute "update 用户 set 银两=银两+" & rnd1 & " where 姓名='" & to1 & "'"
 conn.execute "update 用户 set 操作时间=now() where 姓名='" &aqjh_name& "'"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','奖励','盟主奖励白银"&rnd1&"两')"
 mzl=saysok&"因%%对江湖有功，武林盟主##特奖励%%"&rnd1&"银两！" 
elseif fn1="惩罚" then
 if to_yin<100000 then
  mzl=saysok&"武林盟主##对%%进行罚款失败，%%快成了穷光蛋了，你先放过%%吧！" 
  exit function
 end if
 randomize()
 rnd1=int(rnd*90000)+10000
 conn.execute "update 用户 set 银两=银两-" & rnd1 & " where 姓名='" & to1 & "'"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','盟主罚款白银"&rnd1&"两')"
 mzl=saysok&"%%太不像话了，武林盟主##对%%罚款"&rnd1&"银两！" 
elseif fn1="斩首" then
 conn.execute "update 用户 set 状态='死',事件原因='盟主"&aqjh_name&"斩首' where 姓名='" & to1 &"'"
 mzl= saysok&"武林盟主##喊到:人犯%%斩立决!"
 call boot(to1,"斩立决，操作者："&aqjh_name&","&"盟主暂首")
 conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','斩首',now(),'盟主斩首')"
elseif fn1="踢人" then
 call boot(to1,"踢人，操作者："&aqjh_name&",盟主踢人")
 mzl= saysok&"<bgsound src=wav/gf.wav loop=1>一阵狂风把%%刮出了聊天室(盟主=##)"
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','踢人','盟主踢人')"
else
mzl="<bgsound src=wav/luo.wav loop=3><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chat/f2/img/rightct3.gif></td><td background=../chat/f2/img/rightct4.gif></td><td><img  src=../chat/f2/img/rightct1.gif></td></tr><tr><td background=../chat/f2/img/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr><td valign=center align=center><font style='font-size:9pt;color:red'>========盟主号令========</font></td></tr><tr><td valign=center align=center><font style='font-size:9pt;color:green'>"&fn1&"(##)</td></tr></table></td><td background=../chat/f2/img/rightct08.gif></td></tr><tr><td><img src=../chat/f2/img/rightct9.gif></td><td background=../chat/f2/img/rightct10.gif></td><td><img src=../chat/f2/img/rightct11.gif></td></tr></table></td></table>"
end if
end function
%>