<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'念经
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_nj then
	Response.Write "<script language=JavaScript>{alert('提示：念经需要["&jhdj_nj&"]级才可以操作！');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
mycz=trim(left(says,i-1))
if mycz="/念经" then
	says="<font color=green>【念经】<font color=" & saycolor & ">"+nianjing(towho)+"</font>"
else
	says="<font color=green>【如沐春风】<font color=" & saycolor & ">"+chunfeng(towho)+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'念经
function nianjing(to1)
if to1="大家" or to1=Application("aqjh_automanname") then to1=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if to1<>aqjh_name then
	rs.open "select 门派,道德,通缉 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
	if rs("门派")="天网" or rs("通缉")=True then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("提示：对方为天网杀手，或通缉犯不可以操作!",1)
	end if
	if rs("道德")>10000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("提示：对方道德大于1万不可操作!",1)
	end if
	rs.close
end if
rs.open "select 银两,体力,门派,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("银两")<10000 or rs("体力")<1000 or rs("门派")<>"出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：念经需出家人才可以操作!\n需要1万银两1000体力!",1)
end if
conn.execute "update 用户 set 银两=银两-10000,体力=体力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 道德=道德+150 where 姓名='" & to1 &"'"
if to1<>aqjh_name then
	nianjing="出家人##席地而坐，双手合实，口中念念有词，南…弥…佗…佛……%%道德增长$$redb150$$b点，出家人##消耗体力$$blueb1000$$b点！"
else
	nianjing="出家人##跪倒在佛像面前，虔诚修练，诉说自己在红尘中的罪恶………虔诚的心把佛主打动增加道德$$redb150$$b点……"
end if	
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

'如沐春风
function chunfeng(to1)
to1=trim(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,体力,通缉 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs("门派")="天网" or rs("通缉")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：对方为天网杀手，或通缉犯不可以操作!",1)
end if
if rs("体力")>10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：对方体力大于1万不可操作!",1)
end if
rs.close
rs.open "select 银两,性别,体力,内力,门派,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("性别")="男" then
	if rs("银两")<10000 or rs("体力")<1000 or rs("门派")<>"出家" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("提示：念经需出家人才可以操作!\n男出家人需要1万银两1000体力!",1)
	end if
	conn.execute "update 用户 set 银两=银两-10000,体力=体力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
	chunfeng="男出家人##发现一受伤者(%%)，伸出授助之手，为其治伤。佛光普照霞光万道，%%渐渐恢复了体力$$redb2000$$b点，出家人##消耗体力$$blueb1000$$b点！"
else
	if rs("银两")<1000 or rs("内力")<1000 or rs("门派")<>"出家" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("提示：念经需出家人才可以操作!\n女出家人需要1万银两1000内力!",1)
	end if
	conn.execute "update 用户 set 银两=银两-10000,内力=内力-1000,操作时间=now() where 姓名='" & aqjh_name &"'"
	chunfeng="女出家人##发现一受伤者(%%)，伸出授助之手，用菩提树汁为其治伤，%%渐渐恢复了体力$$redb1500$$b点，出家人##消耗内力$$blueb1000$$b点！"
end if
conn.execute "update 用户 set 体力=体力+2000 where 姓名='" & to1 &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
