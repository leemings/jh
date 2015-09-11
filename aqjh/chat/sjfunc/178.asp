<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'偷窃术
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【偷窃术】<font color=" & saycolor & ">"+tqs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tqs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,道德,grade,体力,智力,职业 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>200 then
Response.Write "<script language=JavaScript>{alert('一次最少偷100最多不能超过200！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="大家" then
 tqs=from1 & "你吃饱了撑着了?!不能向大家使用偷窃术呀！"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用偷窃术最少需要80级！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>200 and rs("grade")<10 then
	Response.Write "<script language=JavaScript>{alert('偷窃术不能大于200！');}</script>"
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
if rs("法力")<25000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够500了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 智力 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs("智力")<fn1 then
tqs=to1&"连喊带叫的哭到：" & aqjh_name & "行行好吧，我那有那么多的智力呀，饶了俺吧,大侠!"
else
conn.execute "update 用户 set 法力=法力-500,体力=体力-"& fn1/1 & ",智力=智力+" & fn1 & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 智力=智力-" & fn1 & " where 姓名='"&to1&"'"
tqs="哎…"& to1 & "真可怜,一觉醒变得白痴白痴的原来被人吸走了" & fn1 & "点智力，难怪看起来有点白痴……(" & aqjh_name & "体力下降"& fn1/1 &")"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>