<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'解毒术♀wWw.happyjh.com♀
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
says="<font color=red>【解毒术】<font color=" & saycolor & ">"+jds(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function jds(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,道德,grade,体力,内力,银两,职业 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 then
Response.Write "<script language=JavaScript>{alert('用的功力太低！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="大家" then
 jds=from1 & "你吃饱了撑着了?!不能向大家使用解毒术呀！"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用解毒术！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用解毒术最少需要80级！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>200000 and rs("grade")<10 then
Response.Write "<script language=JavaScript>{alert('不要太用力了最多20万！');}</script>"
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
if rs("法力")<1000 then
Response.Write "<script language=JavaScript>{alert('你法力不够了至少要1000！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 内力,体力,状态 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs("状态")="正常" then
jds=to1&"看了看：" & sjjh_name & "的兰花指，又看了看自己，我没中毒呀？你要干吗？!"
else
conn.execute "update 用户 set 法力=法力-100,内力=内力-" & fn1 & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 状态='正常' where 姓名='"&to1&"'"
jds=sjjh_name & "使用了" & fn1 & "的内力,解除了潜藏在"& to1 &"体内已久的毒物，"&to1&"感激的直要以身相许。。"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>