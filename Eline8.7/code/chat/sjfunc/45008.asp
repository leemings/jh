<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'布施术♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以布施术！');}</script>"
	Response.End
end if
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
if trim(towho)="" or towho="大家" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>【布施天下】</font><font color=" & saycolor & ">"+bushishu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'布施术
function bushishu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,银两,等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("银两")<5000000 then
Response.Write "<script language=JavaScript>{alert('你身上没有500万，先带500万在身上再说吧！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<5000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得5000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<25 then
Response.Write "<script language=JavaScript>{alert('你此功能需要25级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-5000,道德=道德+1000,银两=银两-5000000 where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两+2000000 where 姓名='" & to1 &"'"
rs.close
rs.open "select 等级,道德 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
ddj=(rs("等级")*1440+2200)+rs("道德")
if rs("道德")>=ddj then
conn.execute "update 用户 set 道德="&ddj&" where 姓名='"&sjjh_name&"'"
end if
bushishu=sjjh_name & "运用法力一式布施天下发给<font color=red>" & to1 & "</font>银两500万两，自己道德增加1000点." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>