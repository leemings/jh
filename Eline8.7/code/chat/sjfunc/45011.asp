<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'魅惑人间♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以魅惑人间！');}</script>"
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
says="<font color=green>【魅惑人间】</font><font color=" & saycolor & ">"+meihuoren(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'魅惑人间
function meihuoren(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,法力,性别,门派 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
dj=rs("等级")
fla=rs("法力")
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('你是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("性别")="男" then
Response.Write "<script language=JavaScript>{alert('此法术只能女孩子使用！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fla<50000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，要50000点才能施展的！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<40 then
Response.Write "<script language=JavaScript>{alert('此功能需要40级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 性别,银两 FROM 用户 WHERE 姓名='" & to1 &"'",conn
xb=rs("性别")
money=int(rs("银两")/4)
if xb="女" then
Response.Write "<script language=JavaScript>{alert('此法术只对男孩子适用！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 银两,门派 FROM 用户 WHERE 姓名='" & to1 &"'",conn
falit=int(rs("银两")/4)
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失败，你不能对高级管理员或站长特封的人员操作!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用户 set 法力=法力-50000 where  姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两+"&money&" WHERE  姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&money&" where 姓名='" & to1 &"'"
meihuoren=sjjh_name & "对" & to1 & "施展<bgsound src=wav/phant030a.wav loop=1>魅惑人间的媚术，" & to1 & "被<img src='img/007.gif'><font color=red>" & sjjh_name & "</font>倾国倾城的容貌所迷惑，身上的银子不知不觉少了四分之一,共计"&money&"两:)..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>