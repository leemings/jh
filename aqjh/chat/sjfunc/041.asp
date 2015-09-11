<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'答题奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用答题奖励！');}</script>"
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
says="<font color=red>【答题奖励】</font><font color=" & saycolor & ">"+zhouyu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'答题奖励
function zhouyu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 知质,法力,智力,等级,门派,职业,金币,武功 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
if rs("金币")<50 then
Response.Write "<script language=JavaScript>{alert('此功能需要50个金币呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("智力")<5 then
Response.Write "<script language=JavaScript>{alert('此功能需要5点智力呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")<>"官府" then
	Response.Write "<script language=JavaScript>{alert('你不是官府人员，不能使用此功能！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 智力=智力-5 where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
zhouyu=to1 & "恭喜你<img src=pic/h9.GIF>，您答对了本江湖举行的答题活动的题目，奖励<font color=red>5</font>个金币，此为随机系统，或多或少，呵呵，谢谢参与活动，祝下次中个大奖!</font>"
 conn.execute "update 用户 set 金币=金币+5 where 姓名='" & to1 &"'" 
case 2 
zhouyu=to1 & "恭喜你<img src=pic/h13.GIF>，您答对了本江湖举行的答题活动的题目，奖励<font color=red>6</font>个金币，此为随机系统，或多或少，呵呵，谢谢参与活动，祝下次中个大奖!</font>"
 conn.execute "update 用户 set 金币=金币+6 where 姓名='"&to1&"'" 
case 3 
zhouyu=to1 & "恭喜你<img src=pic/h21.GIF>，您答对了本江湖举行的答题活动的题目，奖励<font color=red>10</font>个金币，此为随机系统，或多或少，呵呵，谢谢参与活动，祝下次中个大奖!</font>"
 conn.execute "update 用户 set 金币=金币+10 where 姓名='"&to1&"'" 
case 4
zhouyu=to1 & "很遗憾<img src=pic/h37.GIF>，您答对了本江湖举行的答题活动的题目，但抽中<font color=red>1</font>个金币而已，此为随机系统，莫怪莫怪，谢谢参与活动，祝下次中个超级大奖!</font>"
 conn.execute "update 用户 set 金币=金币+1 where 姓名='"&to1&"'" 
case 5
zhouyu=to1 & "恭喜你<img src=pic/h4.GIF>，您答对了本江湖举行的答题活动的题目，抽中一等奖<font color=red>1</font>张会员金卡，此为随机系统，或多或少，呵呵，谢谢参与活动，祝下次中个大奖!</font>"
 conn.execute "update 用户 set 会员金卡=会员金卡+1 where 姓名='"&to1&"'" 
case 6
zhouyu=to1 & "恭喜您<img src=pic/h18.GIF>，您答对了本江湖举行的答题活动的题目的大奖，抽中了头等超级大奖<font color=red>2</font>张会员金卡，您今天运气真好，祝下次中辆劳斯来斯，哈哈!</font>"
 conn.execute "update 用户 set 会员金卡=会员金卡+2 where 姓名='"&to1&"'"
   rs.close	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>