<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'点金术♀wWw.51eline.com♀
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
says="<font color=red>【点金术】<font color=" & saycolor & ">"+djs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function djs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,职业,道德,grade,体力,银两,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
fn1=abs(fn1)
if fn1<100 then
Response.Write "<script language=JavaScript>{alert('要给别人银两也不能少一百吧，别太小气！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if sjjh_name=to1 or to1="大家"  then
 djs=sjjh_name & "你吃饱了撑着了?!不能向大家或自己使用点金术呀！"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("职业")<>"魔法师" then
Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用点金术！！请去职业转换为魔法师！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("等级")<80 then
Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用点金术最少需要80级！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if fn1>50000 and rs("grade")<10 then
Response.Write "<script language=JavaScript>{alert('点金术不能大于五万！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("银两")<fn1 then
Response.Write "<script language=JavaScript>{alert('你哪里那么多的银两？搞错了！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("法力")<500 then
Response.Write "<script language=JavaScript>{alert('你法力不够了要500的法力！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 银两,存款 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs("存款")>50000000 then
djs=to1&"[*^_^*]笑道，俺不是财迷，等我用完了着再给吧，谢谢!"
else
conn.execute "update 用户 set 法力=法力-500,道德=道德+"&fn1/100&",银两=银两-" & fn1 & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 银两=银两+" & fn1*5 & " where 姓名='"&to1&"'"
djs=sjjh_name & "魔法师从身上取出" & fn1 & "两，用魔法棒一指，送给了" & to1 &"，"& to1 &"得到了"& fn1*5&"两白花花的纹银,高兴的连声说谢谢！(道德上升"& fn1/100 &")"
end if

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>
