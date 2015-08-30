<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'配制宝石♀wWw.happyjh.com♀
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
says="<font color=green>【配制宝石】<font color=" & sayscolor & ">"+peibashi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


function peibashi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,w9,等级,法力,操作时间 FROM 用户 WHERE 姓名="&"'" & sjjh_name &"'",conn,3,3
fali=rs("法力")
w6w=rs("w6")
sj=DateDiff("n",rs("操作时间"),now())
if sj<int(rnd*10) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=int(rnd*10)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("等级")<15 then
Response.Write "<script language=JavaScript>{alert('此功能需要15级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

if fali<3000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得3000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"红宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有红宝石？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"绿宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有绿宝石吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"蓝宝石")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你没有蓝宝石吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w6,w9 FROM 用户 WHERE 姓名="&"'" & sjjh_name &"'",conn,3,3
fq=abate(rs("w6"),"蓝宝石",1)
fq=abate(fq,"红宝石",1)
fq=abate(fq,"绿宝石",1)
conn.execute "update 用户 set  w6='"&fq&"' where 姓名='"&sjjh_name&"'"

fq1=add(rs("w9"),"魔力钻石",1)
conn.execute "update 用户 set  w9='"&fq1&"' where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 法力=法力-3000,操作时间=now() where 姓名="&"'" & sjjh_name &"'"
peibashi="##取出红、绿、蓝三种宝石，三种宝石结合在一起，一道光芒升起，<img src='img/look52.gif'>三种宝石化成了魔力钻石." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>