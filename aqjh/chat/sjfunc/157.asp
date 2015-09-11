<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'狼牙棒
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
says="<font color=green>【狼牙棒】<font color=" & sayscolor & ">"+langyabang(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function langyabang(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,门派,法力,等级 FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
w6w=rs("w9")
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('你是出家人不能操作啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
if rs("法力")<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得1000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 门派 FROM 用户 WHERE 姓名="&"'" & towho &"'",conn,3,3
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对出家人操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"狼牙棒")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('你有狼牙棒吗？');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"狼牙棒",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&aqjh_name&"'"
rs.close
rs.open "select * FROM 用户 WHERE 姓名="&"'" & aqjh_name &"'",conn,3,3
conn.execute "update 用户 set 法力=法力-1000 where 姓名="&"'" & aqjh_name &"'"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
application("aqjh_dianxuename")=application("aqjh_dianxuename")&towho&"|"&sj2&"|"&";"
langyabang="##拿着法器<bgsound src=wav/Bombs020.wav loop=1>狼牙棒对着<font color=red>%%</font>就是一阵猛打，不一会儿就把%%打成了白痴，%%真是可怜..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>