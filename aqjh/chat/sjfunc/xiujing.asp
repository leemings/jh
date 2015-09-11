<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'修炼精神
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
says="<font color=red>【种族修炼】<font color=" & saycolor & ">"+dazhuo()+"</font>"
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
if cy<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=20-cy
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<10 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你刚上线才不到10分钟，还是休息会吧！');}</script>"
	Response.End
end if
if rs("知质")<20  then
	Response.Write "<script language=JavaScript>{alert('需20以上的知质才可以！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("智力")<10  then
	Response.Write "<script language=JavaScript>{alert('需10以上的智力才可以！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("种族")<>"人族"  then
	Response.Write "<script language=JavaScript>{alert('此功能为人族才可以使用！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("武功")<rs("等级")*aqjh_nlsx+2000+rs("武功加") then
	conn.execute "update 用户 set 精神=精神+1,知质=知质-20,智力=智力-10,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##作为一个人类，为了提升自身的修为，在深山老林里面顿悟了三天三夜，终于明白了怎么做一个人。自身的精神力<font color=red>+1</font>点!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("武功加")<=rs("等级")*15000 then
	conn.execute "update 用户 set 精神=精神+2,知质=知质-20,智力=智力-10,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##为了寻求人类的精华,自觉在深山里修炼，不顾风吹雨打，精神力提高<font color=red>2</font>点,为以后的修为打下了基础，真是一个上进的青年!"
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