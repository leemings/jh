<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'使用♀wWw.happyjh.com♀
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【使用药品】<font color=" & saycolor & ">"+eat(mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'使用
function eat(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：物品名|数量 ');}</script>"
	Response.End 
end if
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
'房间处理
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="高手房间" and Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<=22) then
	if instr(fn1,"千年人参")=0 or instr(fn1,"万年灵芝")=0 then
		Response.Write "<script language=JavaScript>{alert('提示：夺宝期间只可以服用[千年人参]及[万年灵芝]两种药品来补充体力内力！');}</script>"
		Response.End
	end if
else
	if instr(fn1,"千年人参")<>0 or instr(fn1,"万年灵芝")<>0 then
		Response.Write "<script language=JavaScript>{alert('提示：只有在夺宝大赛时才可以使用[千年人参]及[万年灵芝]两种药品！');}</script>"
		Response.End
	end if
end if
zswupin=trim(zt(0))
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于100！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
eattemp=abate(rs("w1"),zswupin,wusl)
rs.close
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
conn.execute "update 用户 set w1='"&eattemp&"',内力=内力+"&nl&",体力=体力+"&tl&" where  姓名='"&sjjh_name&"'"
eat="##使用了药品[<b>"&zswupin&"</b>]共计："&wusl&"个,自己的体力：<font color=red>+"&tl&"</font>点,内力:<font color=red>+"&nl&"</font>点!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>