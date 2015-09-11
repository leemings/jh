<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'神灵降世
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=fakuan(mid(says,i+1)+0,towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'神灵降世"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 师傅,智力,银两,杀人数,总杀人,sl,会员金卡 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>小神棍</font>]对[<font color=red>"&aqjh_name&"</font>]说：你哪里那么多的钱？搞错了，你比我还神棍啊！"
	exit function
end if
if rs("sl")<>" " and rs("sl")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
      fakuan="<font color=red>【神灵降世】</font>友情提示：[<font color=red>"&aqjh_name&"</font>]您已经有神灵帮助了，不要再来了，金卡多就送给新人吧！"
	exit function
end if
if rs("会员金卡")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>大神棍</font>]对[<font color=red>"&aqjh_name&"</font>]说：你哪里有<img src='picwords/1.gif'><img src='picwords/0.gif'>块金卡？搞错了，你比我还神棍啊！"
	exit function
end if
rs.close
rs.open "SELECT 银两,grade,师傅,智力,总杀人,sl,slsj,会员金卡,知质,智力 FROM 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("总杀人")>20 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan=""&aqjh_name&"想祈祷神灵[封神]降临辅助，但由于"&aqjh_name&"心术不良，总杀人数超过<img src='picwords/2.gif'><img src='picwords/0.gif'>人！请不到[封神]降临帮忙！"
	exit function
end if
if rs("银两")>=fn1 then
    conn.execute("update 用户 set 知质=知质+10,智力=智力+10,sl='封神',slsj=now()+2,会员金卡=会员金卡-10,银两=银两-" & fn1 & " where 姓名='"&aqjh_name&"'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【神灵降世】</font>"&aqjh_name&"向一江湖神棍买了一张[<font color=red>神灵符</font>]，花了" & fn1 & "两，半信半疑的用<font color=red>10</font>块金卡祭神，哇~[<b><font color=red>封神</font></b>]真的降临"&aqjh_name&"身上，泡点速度增加2倍，时间为1天，神灵附身后智力上升10点，知质上升10点！"
	exit function
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
