<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'NPC召唤术
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
says="<font color=green>【NPC召唤】<font color=" & saycolor & ">"+zhnpc(mid(says,i+1),towho)+"</font>"
call chatsay("NPC召唤",towhoway,towho,saycolor,addwordcolor,addsays,says)
function zhnpc(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'开始判断npc
rs.open "select * from npc where n姓名='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：NPC["&to1&"]并不存在或者它不是NPC你搞错了吧！');}</script>"
	Response.End	
end if
if InStr(";" & Application("aqjh_npc"), ";" &to1& "|")=0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：NPC["&to1&"]不在线！');}</script>"
	Response.End
end if
if rs("n敌人")=aqjh_name then
  zhnpc="##:你是NPC[%%]的敌人，它不会跟你一起走的，你走开吧！"
  exit function
end if
n_zhuren=rs("n主人")
rs.close
'开始判断npc主人有没有死
if n_zhuren<>"无" then
rs.open "select * from 用户 where 姓名='" &n_zhuren&"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：主人["&z_zhuren&"]不存在,NPC出现异常！');}</script>"
	Response.End	
end if
if rs("状态")<>"正常" then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：NPC["&to1&"]的主人["&z_zhuren&"]还没有死呢！');}</script>"
	Response.End
end if
elseif n_zhuren="无" then
 	Response.Write "<script language=JavaScript>{alert('提示：官府发放的NPC，不允许召唤！');}</script>"
	Response.End
end if
if n_zhuren=aqjh_name then
   zhnpc="提示：##:你搞什么啊，你已经是NPC[%%]的人主了！"
   exit function
end if
'判断自己
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("体力")<10000 or rs("内力")<10000 then
   zhnpc="##想对%%使用[召唤术]，嘛迷嘛迷轰....无奈自己道行太浅此次操作失败了,白白消耗原气!"
   exit function
end if
if rs("金币")<5 then
   zhnpc="提示：##:金币没有5个，NPC不跟你走！"
   exit function
end if
if rs("职业")<>"魔法师" then
   zhnpc="提示：##:你不是魔法师，你没有召唤的功能！"
   exit function
end if
conn.execute "update npc set n敌人='无',n主人='"&aqjh_name&"' where n姓名='"&to1&"'"
conn.execute "update 用户 set 体力=体力-10000,内力=内力-10000,金币=金币-5 where 姓名='"&aqjh_name&"'"
zhnpc="[##]对NPC人物[%%]施展[召唤术]：乖乖听话，跟我混一定没错的，来吧，来吧......(NPC人物[%%]失去了心志，##成了%%的主人)"
rs.close
conn.close
end function
%>