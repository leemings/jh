<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'哑穴♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('想作什么呀，你的管理等级可不够呀！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="高手E线" then
	Response.Write "<script language=JavaScript>{alert('提示：在夺宝大赛期间不可以使用点别人哑穴！');}</script>"
	Response.End
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【哑穴】<font color=" & saycolor & ">"+ya(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'哑穴
function ya(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
fn1=Replace(fn1,"=","")
fn1=Replace(fn1,"untion","")
fn1=Replace(fn1,chr(39),"")
rs.open "select grade,门派 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
grade=rs("grade")
menpai=rs("门派")
denji=rs("grade")
rs.close
rs.open "select 门派,grade from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没有这个人？你是不是看错了！');}</script>"
	Response.End
end if
if denji<6 and menpai<>rs("门派") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：搞错了["&to1&"]也不是你们门派的呀！');}</script>"
	Response.End
end if
if rs("grade")<grade  then
	Application.Lock
	application("sjjh_dianxuename")=application("sjjh_dianxuename")&to1&"|"&now()&"|"&";"
	Application.UnLock
	ya="##对%%使用了哑穴术，" & to1 & "呆呆地不动了……"
else
	ya="##对%%使用了哑穴术，可是你的等级不如人家？没办法！"
end if
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','哑穴','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>