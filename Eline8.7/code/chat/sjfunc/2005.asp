<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'平安术♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【平安术】<font color=" & saycolor & ">"+pas(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function pas(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,道德,职业,grade,体力,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*2)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*2)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
fn1=abs(fn1)
if fn1<100 or fn1>50000 then
Response.Write "<script language=JavaScript>{alert('要使用平安术最少一百最多不能超过5万！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="大家" then
 pas=sjjh_name & "你吃饱了撑着了?!不能向大家使用平安术呀！"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用平安术！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用平安术最少需要80级！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if fn1>50000 and rs("grade")<10 then
	Response.Write "<script language=JavaScript>{alert('平安术不能大于五万！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("体力")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪里那么多的体力？搞错了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("法力")<500 then
	Response.Write "<script language=JavaScript>{alert('你法力不够了，要500的法力！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,轻功,法力 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs("体力")>300000 then
pas=to1&"笑嘻嘻道：多谢" & sjjh_name & "，我现在身体还壮着呢,多谢你关心!我不需要平安术"
else
conn.execute "update 用户 set 法力=法力-500,轻功=轻功+"&fn1/100&",体力=体力-" & fn1 & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 体力=体力+"&fn1*10&" where 姓名='"&to1&"'"
pas=sjjh_name & "双手轻轻一挥将" & fn1 & "的体力传入了" & to1 &"体内，"& to1 &"体力和体力上限猛增"& fn1*10&",高兴的眼都睁不开了[*^_^*]！(" & sjjh_name & "轻功上升"& fn1/100 &")"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>