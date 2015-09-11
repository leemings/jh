<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'魔界咒语
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用魔界咒语！');}</script>"
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
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="大家" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>【魔法师·魔界咒语】</font><font color=" & saycolor & ">"+zhouyu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'魔界咒语
function zhouyu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,门派,职业 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('你是出家人啊啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="天网"  then
Response.Write "<script language=JavaScript>{alert('天网的杀手来这里做什么!');}</script>"
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
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢！！请去职业转换为魔法师！');}</script>"
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
if rs("法力")<50000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得50000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
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
conn.execute "update 用户 set 法力=法力-3000 where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 内力=内力-8000 where 姓名='" & to1 &"'"
zhouyu=aqjh_name & "运用法力狂唸咒语，<font color=red>" & to1 & "</font>内力逐渐消失,真是死定了..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>t 内力=内力-8000 where 姓名='" & to1 &"'"
zhouyu=aqjh_name & "运用法力狂唸咒语，<font color=red>" & to1 & "</font>内力逐渐消失,真是死定了..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>