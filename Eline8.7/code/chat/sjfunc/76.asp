<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'出家/还俗♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(0)="高手E线" then
	Response.Write "<script language=JavaScript>{alert('提示：要出家或还俗到别的房间去，夺宝大赛中不可以出家！');}</script>"
	Response.End
end if
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="E线江湖"  then
Response.Write "<script Language=Javascript>alert('提示：[E线江湖]房间里现在是只给堂主和护法，长老，掌门等进行门派大战，其他人等在场可让你门派加强，想打架到[高手E线]房间去吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
erase sjjh_roominfo
erase chatinfo
if sjjh_jhdj<sjjh_jhcz then
	Response.Write "<script language=JavaScript>{alert('提示：想出家也要["&sjjh_jhcz&"]级才行！');}</script>"
	Response.End
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
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/出家" then
	says="<font color=green>【看破红尘】</font><font color=" & saycolor & ">"+chujia()+"</font>"
else
	says="<font color=green>【还俗】</font><font color=" & saycolor & ">"+huanshu()+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function chujia()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")

rs.open "select grade,身份,配偶,门派,情人 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是管理员不可以操作，除非你不想作了!');}</script>"
	Response.End
end if
if rs("门派")="天网" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：天网不可以出家!');}</script>"
	Response.End
end if
if rs("身份")="掌门" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是掌门不可以离开自己的门派');}</script>"
	Response.End
end if
if rs("配偶")<>"无" or rs("情人")<>"无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你红尘心事未了,不能出家\n(不可有配偶及情人)！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 门派='出家',身份='弟子',grade=1,保护=True where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("出家")
chujia="##看破红尘,不再有任何的眷恋,决定今日落发出家,以后江湖上的打打杀杀,江湖恩怨与我无关~~"
end function

'还俗
function huanshu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 门派 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("门派")<>"出家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你非出家人,不可以操作!');}</script>"
	Response.End
end if
conn.execute "update 用户 set 门派='游侠',身份='弟子',grade=1,保护=True,存款=int(存款/2),银两=int(银两/2) where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("游侠")
huanshu="##红尘心事为了,整天不一心理佛,三心二意,最终还是决定离开此地,走时把自己身上钱财留下一半给了寺中兄弟~~"
end function

sub updatemd(jhmp)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select id,内力,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
sjjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
jhtx=rs("名单头像")
sex=rs("性别")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=sjjh_name then
		sjjh_zm=sjjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&sjjh_jhdj&"|"&sjjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)
		onlinelist(i)=sjjh_zm
		exit for
	end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end sub
%>