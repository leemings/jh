<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'百变神通♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")

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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
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
conn.open Application("sjjh_usermdb")
rs.open "select w6,法力,操作时间 FROM 用户 WHERE 姓名="&"'" & sjjh_name &"'",conn,3,3
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
fla=rs("法力")
if fla<200 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得200点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-200,操作时间=now() where 姓名="&"'" & sjjh_name &"'"
randomize()
'r=int(rnd()*3)
r=int(rnd*7)+4

select case r
case 1


	zstemp=add(rs("w6"),"冰水",1)
	conn.execute "update 用户 set  w6='"&zstemp&"' where 姓名='"&sjjh_name&"'"
	rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个冰水，消耗法力"&jhfl/10&"."

case 2


zstemp=add(rs("w6"),"矿石",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & sjjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个矿石，消耗法力"&jhfl/10&"."

case 3


zstemp=add(rs("w6"),"大鲨鱼",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & sjjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个大鲨鱼，消耗法力"&jhfl/10&"."


case 4

zstemp=add(rs("w6"),"大草鱼",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & sjjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个大草鱼，消耗法力"&jhfl/10&"."
case 5

zstemp=add(rs("w6"),"小鲤鱼",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & sjjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一个小鲤鱼，消耗法力"&jhfl/10&"."

case 6

zstemp=add(rs("w6"),"老虎肉",1)
conn.execute "update 用户 set w6='"&zstemp&"' where 姓名="&"'" & sjjh_name &"'"
rs.close
banbianshu="##你施展法术<bgsound src=wav/Phant008.wav loop=1>变出了一块老虎肉，消耗法力"&jhfl/10&"."

case else
banbianshu="##你<bgsound src=wav/Phant008.wav loop=1>运气真是差差呀，什么都没变出来."
rs.close
end select

end function 
%>