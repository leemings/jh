<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'迷魂大法♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以迷魂大法！');}</script>"
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
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="大家" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>【迷魂大法】</font><font color=" & saycolor & ">"+mihunfa(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'迷魂大法
function mihunfa(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,门派,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
sj=DateDiff("s",rs("操作时间"),now())
if sj<9 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=9-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<30 then
Response.Write "<script language=JavaScript>{alert('此功能需要30级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<30000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得30000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 法力,门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
falit=int(rs("法力")/2)
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-30000+"&falit&" where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 法力=法力-"&falit&" where 姓名='" & to1 &"'"
mihunfa=sjjh_name & "对" & to1 & "施展<bgsound src=wav/fx40a.wav loop=1>迷魂大法，" & to1 & "法力逐渐被<font color=red>" & sjjh_name & "</font>所吸收掉"&falit&"点，" & to1 & "呆呆地在那傻笑着，呼..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>