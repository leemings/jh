<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'教武chiwu
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
says="<font color=red>【指导武功】<font color=" & saycolor & ">"+chiwu()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'教武chiwu
function chiwu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 内力,职业,体力,等级,道德,武功,武功加,操作时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
if rs("职业")<>"武师" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的职业不是武师，不能操作！');}</script>"
	Response.End
end if

sj=DateDiff("s",rs("操作时间"),now())
if sj<600 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=600-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("内力")<5000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要内力5000才可以操作！');}</script>"
	Response.End
end if
if rs("武功")<rs("等级")*aqjh_wgsx+3800+rs("武功加") then
	conn.execute "update 用户 set 武功=武功+1000,道德=道德-200,操作时间=now() where 姓名='" & aqjh_name &"'"
	chiwu="武师""##<bgsound src=wav/dz.wav loop=1>为了提升自己的武功，不惜冒着走火入魔的危险进入火云洞<font color=red>终于成功恢复武功</font>1000点,道德下降<font color=red>+3500</font>100点,真是一山要比一山高!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("武功加")<=rs("等级")*1500 then
	conn.execute "update 用户 set 武功加=武功加+5000,道德=道德-500,操作时间=now() where 姓名='" & aqjh_name &"'"
	chiwu="武师""##<bgsound src=wav/dz.wav loop=1>为了提升自己的武功，不惜冒着走火入魔的危险进入火云洞<font color=red>终于成功恢复武功</font>5000点,道德下降<font color=red>-500</font>点,真是一山要比一山高!!"
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：现在你的上限满了，等升了级再指导武功吧吧');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>