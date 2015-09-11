<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'布施术
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以布施术！');}</script>"
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
says="<font color=green>【魔法师·布施天下】</font><font color=" & saycolor & ">"+bushishu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'布施术
function bushishu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,银两,职业,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<3000000 then
Response.Write "<script language=JavaScript>{alert('你身上没有300万，先带300万在身上再说吧！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等"& s &"秒再操作,可别累着！');</script>"
	response.end
end if
if rs("法力")<2000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得2000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-2000,道德=道德+3000,银两=银两-3000000,操作时间=now() where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 银两=银两+2000000 where 姓名='" & to1 &"'"
rs.close
rs.open "select 等级,道德 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
ddj=(rs("等级")*1440+2200)+rs("道德")
if rs("道德")>=ddj then
conn.execute "update 用户 set 道德="&ddj&" where 姓名='"&aqjh_name&"'"
end if
bushishu=aqjh_name & "运用法力一式布施天下发给<font color=red>" & to1 & "</font>银两200万两，自己损失银两300万两，法力下降1000点，自己道德增加3000点." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>f'><img src='picwords/0.gif'>万两，法力下降<img src='picwords/1.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>点，自己道德增加<img src='picwords/3.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>点." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>