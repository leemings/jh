<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'照顾新人
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
'照顾新人"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 姓名='" & to1 & "'",conn,2,2
if rs("体力加")>50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【照顾新人】</font><font color=blue>##</font>，你搞错了吧，<font color=blue>%%</font>他已经被你照顾一次了，不能再来了！</font>"
	exit function
end if
if rs("体力")>1000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【照顾新人】</font><font color=blue>##</font>，你搞错了吧，<font color=blue>%%</font>他并不是新人啊！</font>"
	exit function
end if
if rs("等级")>18 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【照顾新人】</font><font color=blue>##</font>，你搞错了吧，<font color=blue>%%</font>他并不是新人啊！</font>"
	exit function
end if
if fn1>5000 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & "><font color=red>【照顾新人】</font><font color=blue>##</font>照顾快乐江湖新人，请来得到高僧及快乐大红人传授，##要求传授<font color=red>" & fn1 & "</font>奖励新人<font color=blue>[%%]</font>，但由于数目太多，失败，二位高人答应的最多只有<font color=red>5000</font>点！</font>"
	exit function
end if
if rs("银两")>=fn1 then
    conn.execute("update 用户 set 体力加=体力加+100000,体力=体力+1000000,魅力=魅力+" & fn1 & " where 姓名='" & to1 & "'")
    conn.execute("update 用户 set 道德=道德+1000 where 姓名='" & aqjh_name & "'")
  	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>【照顾新人】</font><font color=blue>[##]</font>遵照快乐良风，奉站长之命帮助新人<font color=blue>[%%]</font>，<font color=blue>[%%]</font>体力增加<font color=red>100万</font>，体力上限增加<font color=red>10万</font>，魅力增加了<font color=red>" & fn1 & "</font>，<font color=blue>[##]</font>由于帮助新人，道德上涨<font color=red>1000</font>点，还多了一些莫名奇妙的属性！</font>"
	exit function
end if
fakuan="<font color=blue><font color=red>##</font>,<font color=red>%%</font>他已经有<font color=red>"&fn1&"</font>。。。。。。。。</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
