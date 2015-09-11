<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'出家/还俗
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<aqjh_jhcz then
	Response.Write "<script language=JavaScript>{alert('提示：想出家也要["&aqjh_jhcz&"]级才行！');}</script>"
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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/出家" then
	says="<font color=green>【遁入空门】</font><font color=" & saycolor & ">"+chujia()+"</font>"
else
	says="<font color=green>【还俗入世】</font><font color=" & saycolor & ">"+huanshu()+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function chujia()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select grade,身份,配偶,门派,总杀人 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是管理员不可以操作，除非你不想作了!');}</script>"
	Response.End
end if
if rs("身份")="掌门" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是掌门还要出家么');}</script>"
	Response.End
end if
if rs("总杀人")>15 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，佛门不收留你这种人');}</script>"
	Response.End
end if
if rs("门派")="天网" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是天网杀手就不用出家了');}</script>"
	Response.End
end if
conn.execute "update 用户 set 门派='出家',身份='和尚',grade=1,保护=True where 姓名='" & aqjh_name &"'"
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
conn.open Application("aqjh_usermdb")
rs.open "select 门派 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("门派")<>"出家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你非出家人,不可以操作!');}</script>"
	Response.End
end if
conn.execute "update 用户 set 离派=离派+1,门派='游侠',身份='弟子',grade=1,保护=True,存款=int(存款/2),银两=int(银两/2) where 姓名='" & aqjh_name &"'"
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
conn.open Application("aqjh_usermdb")
rs.open "select id,内力,操作时间,会员等级,身份,门派,名单头像,性别,好友名单,配偶 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
jhtx=rs("名单头像")
sex=rs("性别")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=aqjh_name then
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end sub
%>