<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'提点智力
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
says="<font color=red>【智力修习】<font color=" & saycolor & ">"+dazhuo()+"</font>"
towhoway=0
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'打坐
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
cy=DateDiff("n",rs("操作时间"),now())
if cy<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-cy
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<4 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你刚上线才不到4分钟，还是休息会吧！');}</script>"
	Response.End
end if
if rs("体力")<20000  then
	Response.Write "<script language=JavaScript>{alert('需20000以上的体力才可以！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("体力")<rs("等级")*aqjh_nlsx+2000+rs("体力加") then
	conn.execute "update 用户 set 智力=智力+60,体力=体力-10000,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>在快乐江湖智力院向智力哒摩拜师，经提点后有所领悟<bgsound src=wav/dz.wav loop=1>,疲劳增加使得体力失去-10000，智力提升+60,从今世上将多了一位有智之人，希望能把智慧发挥正途上!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("体力加")<=rs("等级")*15000 then
	conn.execute "update 用户 set 智力=智力+80,体力=体力-10000,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>在快乐江湖智力院向智力哒摩拜师，经提点后有所领悟<bgsound src=wav/dz.wav loop=1>,疲劳增加使得体力失去<font color=red>-10000</font>点,智力提升提升<font color=red>+80</font>点,一代功名万古枯!"
else
	Response.Write "<script language=JavaScript>{alert('现在你的上限满了，等升了级再练吧');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>