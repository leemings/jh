<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'奖励徒弟
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
'奖励徒弟"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 师傅,智力,银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪里那么多的钱？搞错了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "SELECT 银两,grade,师傅,智力 FROM 用户 where 姓名='" & to1 & "' and 师傅='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[官府]对[##]说：[%%]不是你的笨徒弟啊！你不能奖励[%%]"
	exit function
end if
if fn1>200000 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="##想奖励徒弟%%银两" & fn1 & "，但由于数目超过江湖规定,！快乐规定奖励最大限制为20万！"
	exit function
end if
if rs("银两")>=fn1 then
    conn.execute("update 用户 set 智力=智力+2,银两=银两+" & fn1 & " where 姓名='" & to1 & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=green>【奖励徒弟】</font>##拍着徒弟</font>%%<font color=red>的脑门说:乖徒儿，师傅给你银子,</font>" & fn1 & "<font color=red>去买棉花糖，以后可要乖乖的啊!%%高兴的眼都睁不开了[*^_^*]</font>（徒弟智力加2）"
	exit function
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
