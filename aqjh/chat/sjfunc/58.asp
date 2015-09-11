<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'本派刑法
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('本派刑法需要4级才可以操作！');}</script>"
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
'本派刑法"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 门派,身份 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sf=rs("身份")
menpai=rs("门派")
if  sf<>"掌门" and sf<>"长老" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="[江湖]对[##]说：你没有权限使用本功能！"
    exit function
end if
rs.close
rs.open "SELECT 内力,grade FROM 用户 where 姓名='" & to1 & "' and 门派='" & menpai & "'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[官符]对[##]说：[%%]不是你门派的人啊！"
	exit function
end if
if aqjh_grade<=rs("grade") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[官符]对[##]说：[%%]是你派的掌门或者是长老，你是不能对他操作的！"
	exit function
end if
if fn1>10000 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="##想罚%%内力" & fn1 & "，但由于数目超过江湖规定,罚款失败！江湖规定刑法最大限制为1万！"
	exit function
end if
if rs("内力")>=fn1 then
    conn.execute("update 用户 set 内力=内力-" & fn1 & " where 姓名='" & to1 & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=green>【本派刑法】</font>##认为</font></b>本门弟子%%<font color=red><b>确实不是什么好人，绝定刑法减少</b></font>" & fn1 & "<font color=red><b>内力，</b></font>"
	exit function
end if
fakuan="<font color=red>##</font><font color=#000000><b>你还是就放过</font></b><font color=red>%%</font><font color=#000000><b>吧，他的身上总共都没有" & fn1&"内力</b></font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
