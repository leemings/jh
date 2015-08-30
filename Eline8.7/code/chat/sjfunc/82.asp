<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'加入离开天网♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<jhdj_jhtw then
	Response.Write "<script language=JavaScript>{alert('提示：想加入天网也要["&jhdj_jhtw&"]级才行！');}</script>"
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
if mycz="/加入天网" then
	says="<font color=green>【加入天网】</font><font color=" & saycolor & ">"+jiarutw()+"</font>"
else
	says="<font color=green>【离开天网】</font><font color=" & saycolor & ">"+likaitw()+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'加入天网
function jiarutw()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select grade,银两,身份 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
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
	Response.Write "<script language=JavaScript>{alert('提示：你是掌门不可以离开自己的门派');}</script>"
	Response.End
end if
if rs("银两")<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：加入天网需要银两大于5000万');}</script>"
	Response.End
end if
conn.execute "update 用户 set 门派='天网',身份='杀手',grade=1,保护=False where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("天网")
jiarutw="##决心加入天网，作天下第一杀手！$$blueb注：加入天网后不可以保护，可以接受别人的杀人请求，成为杀手!$$b"
end function

'离开天网
function likaitw()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 门派 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("门派")<>"天网" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你非天网杀手,不可以操作!');}</script>"
	Response.End
end if
conn.execute "update 用户 set 门派='游侠',身份='弟子',grade=1,保护=True,存款=int(存款/2),银两=int(银两/2) where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("游侠")
likaitw="##已经厌倦了打打杀杀的生活，不再为天网效力。$$red拿出自己积蓄的一半给了组织，才申请离开了天网!$$"
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