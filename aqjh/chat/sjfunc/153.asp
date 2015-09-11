<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'抢劫令
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
says="<font color=green>【抢劫令】<font color=" & sayscolor & ">"+qiangjielin(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function qiangjielin(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,门派,法力,等级 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
w6w=rs("w9")
if rs("门派")="官府" and aqjh_grade<10 then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你是出家人想做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('所有有动武倾向的都在每小时的前25分钟，现在是聊天时间！');window.close();}</script>"
	Response.End 
end if
if rs("等级")<60 then
Response.Write "<script language=JavaScript>{alert('此功能需要60级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<8000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得8000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 门派,银两 FROM 用户 WHERE 姓名="&"'" & towho &"'",conn,3,3
money=int(rs("银两")/2)
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你想对出家人做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
'rs.close
wuyn=iswp(w6w,"抢劫令")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你有抢劫令吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"抢劫令",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&aqjh_name&"'"

conn.execute "update 用户 set 法力=法力-8000 where 姓名="&"'" & aqjh_name &"'"
conn.execute "update 用户 set 银两=银两+"&money&" WHERE  姓名="&"'" & aqjh_name &"'"
conn.execute "update 用户 set 银两=银两-"&money&" where 姓名="&"'" & towho &"'"
qiangjielin="##拿出法器<bgsound src=wav/Phant012.wav loop=1>抢劫令,对着<font color=red>%%</font>大吼一声:抢劫……,从%%那抢得银两"&money&"两,##消耗法力8000点" 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>