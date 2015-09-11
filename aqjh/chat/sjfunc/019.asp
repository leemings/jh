<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'惩罚徒弟
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
'惩罚徒弟"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 师傅,道德,银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("道德")<2000 then
	Response.Write "<script language=JavaScript>{alert('你自己没道德了，你知道你很差吗？你还有什么资格惩罚你徒弟！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "SELECT 道德,grade,师傅,魅力 FROM 用户 where 姓名='" & to1 & "' and 师傅='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[官府]对[##]说：[%%]不是你的徒弟啊！你不能惩罚[%%]，当心他去告你哦，让你吃不了兜着走！"
	exit function
end if
if fn1>10000 then
    conn.execute("update 用户 set 道德=道德-2000 where 姓名='" & aqjh_name & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="##想对徒弟%%惩罚他" & fn1 & "点魅力，但由于师傅##太黑心,自己减少了2000点道德！这难道是为人师表吗！"
	exit function
end if
if rs("魅力")>=fn1 then
    conn.execute("update 用户 set 魅力=魅力-" & fn1 & " where 姓名='" & to1 & "'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=green>【惩罚徒弟】</font>##拍着徒弟</font>%%<font color=red>的脑门说:好你个徒儿，敢不听话,做出危害师门，危害江湖的事来，我罚你</font>" & fn1 & "<font color=red>魅力，让你没人喜欢，讨人厌!</font>"
	exit function
end if
fakuan="<font color=red>##</font><font color=red>你还是就放过你徒弟</font><font color=red>%%</font><font color=red>吧，他已经很没人缘了，连" & fn1&"点魅力都没，给别人留个好印象好不？？</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>