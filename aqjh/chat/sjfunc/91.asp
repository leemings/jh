<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'没收魔器
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
says="<font color=green>【没收魔器】<font color=" & sayscolor & ">"+mushou(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function mushou(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3

if rs("门派")<>"官府" then
Response.Write "<script language=JavaScript>{alert('此功能官府人员方可使用呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w9 FROM 用户 WHERE 姓名="&"'" & towho &"'",conn,3,3
w6w=rs("w9")
wuyn=iswp(w6w,"狼牙棒")
if wuyn=1 then
fq=abate(w6w,"狼牙棒",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&towho&"'"
yn=1
end if
wuyn=iswp(w6w,"破天锥")
if wuyn=1 then
fq=abate(w6w,"破天锥",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&towho&"'"
yn=1
end if
wuyn=iswp(w6w,"血滴子")
if wuyn=1 then
fq=abate(w6w,"血滴子",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&towho&"'"
yn=1
end if
wuyn=iswp(w6w,"抢劫令")
if wuyn=1 then
fq=abate(w6w,"抢劫令",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&towho&"'"
yn=1
end if
wuyn=iswp(w6w,"绝情刀")
if wuyn=1 then
fq=abate(w6w,"绝情刀",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&towho&"'"
yn=1
end if
if yn=0  then
Response.Write "<script language=JavaScript>{alert('对方什么魔器都没有啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
mushou="##一声大喝：<font color=red>%%</font>，谁叫你在非动武时间乱用魔器伤人的，没有王法了吗，惩戒一下收掉你的法宝." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>