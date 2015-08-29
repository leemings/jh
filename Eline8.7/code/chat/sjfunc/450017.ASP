<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'发射子弹♀wWw.51eline.com♀
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
says="<font color=green>【发射子弹】<font color=" & sayscolor & ">"+fashezi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


function fashezi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,门派,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
fali=rs("法力")
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<20 then
Response.Write "<script language=JavaScript>{alert('此功能需要20级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 体力,门派 FROM 用户 WHERE 姓名='" & towho &"'",conn
money=int(rs("体力")/10)
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w6 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
w6w=rs("w6")

wuyn=iswp(w6w,"手枪")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你有精制手枪吗？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
wuyn=iswp(w6w,"子弹")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你有子弹吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fali<8000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得8000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"子弹",1)
conn.execute "update 用户 set  w6='"&fq&"' where 姓名='"&sjjh_name&"'"

conn.execute "update 用户 set 法力=法力-5000 where 姓名="&"'" & sjjh_name &"'"

randomize()

'rnd1=int(rnd()*3)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update 用户 set 体力=体力-"&money&" where 姓名='" & towho &"'"
fashezi="##从腰间拔出一把精制手枪<bgsound src=wav/Phant012.wav loop=1>，恶狠狠地对着<font color=red>%%</font>就是一枪,%%啊的一声，体力失去1/10，共计"&money&"点......" 
rs.close
rs.open "select 体力 from 用户 where 姓名='" & towho &"'",conn

if rs("体力")<=-100 then 
conn.execute "update 用户 set 状态='死' where 姓名='" & towho &"'"
fashezi="##从腰间拔出一把精制手枪<bgsound src=wav/Phant012.wav loop=1>，恶狠狠地对着<font color=red>%%</font>就是一枪,%%啊的一声，由于体力不支，被当场击毙，好惨啊..." 
  call boot(towho,"枪杀，操作者："&sjjh_name&","&fn1)
end if


case else
fashezi="##从腰间拔出一把精制手枪<bgsound src=wav/Phant012.wav loop=1>，恶狠狠地对着<font color=red>%%</font>就是一枪,怎奈枪法太臭没打中......"
end select

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>