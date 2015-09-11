<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'打坐
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【打坐】<font color=" & saycolor & ">"+dazhuo()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'打坐
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 体力,等级,内力加,内力,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("体力")<2100  then
	Response.Write "<script language=JavaScript>{alert('需2000以上的体力才可以！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("内力")<rs("等级")*aqjh_nlsx+2000+rs("内力加") then
	conn.execute "update 用户 set 内力=内力+300,体力=体力-2000,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>进行练功打坐,体力失去-2000内力提升+300,有得必有失嘛!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("内力加")<=rs("等级")*1500 then
	conn.execute "update 用户 set 内力加=内力加+9,体力=体力-2000,操作时间=now() where 姓名='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>进行练功打坐,体力失去<font color=red>-2000</font>点,内力上限提升<font color=red>+9</font>点,有得必有失嘛!"
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