<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'百变神通
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says="<font color=green>【百变神通】<font color=" & sayscolor & ">"+banbianshu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function banbianshu(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,法力,知质,职业 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
fla=rs("法力")
if fla<2000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得2000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("职业")<>"炼金师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢！！请去职业转换为炼金师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-2000 where 姓名="&"'" & aqjh_name &"'"
randomize()
'r=int(rnd()*3)
r=int(rnd*7)+4
select case r
case 1
	zstemp=add(rs("w6"),"冰水",1)
	conn.execute "update 用户 set 知质=知质+2,w6='"&zstemp&"' where 姓名='"&aqjh_name&"'"
	rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个冰水，消耗法力200点，知质增加2点."
case 2
zstemp=add(rs("w6"),"矿石",1)
conn.execute "update 用户 set 知质=知质+2,w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个矿石，消耗法力200点，知质增加2点."
case 3
zstemp=add(rs("w6"),"大鲨鱼",1)
conn.execute "update 用户 set 知质=知质+2,w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个大鲨鱼，消耗法力200点，知质增加2点."
case 4
zstemp=add(rs("w6"),"大草鱼",1)
conn.execute "update 用户 set 知质=知质+2,w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个大草鱼，消耗法力200点，知质增加2点."
case 5
zstemp=add(rs("w6"),"小鲤鱼",1)
conn.execute "update 用户 set 知质=知质+10,w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个小鲤鱼，消耗法力200点，知质增加10点."
case 6
zstemp=add(rs("w6"),"老虎肉",1)
conn.execute "update 用户 set 知质=知质+2,w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一块老虎肉，消耗法力200点，知质增加2点."
case else
banbianshu="##你<bgsound src=wav/Phant008.wav loop=1>运气真是差差呀，什么都没变出来."
rs.close
end select
end function 
%>了<img src='picwords/1.gif'>个小鲤鱼，消耗法力<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>点."
case 6
zstemp=add(rs("w6"),"老虎肉",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & aqjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了<img src='picwords/1.gif'>块老虎肉，消耗法力<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>点."
case else
banbianshu="##你<bgsound src=wav/Phant008.wav loop=1>运气真是差差呀，什么都没变出来."
rs.close
end select
end function 
%>