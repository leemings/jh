<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'追杀令
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('追杀令需要4级才可以操作！');}</script>"
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
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=ling(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'追杀令
function ling(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 身份,银两,门派 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,2,2
sf=rs("身份")
menpai=rs("门派")
if sf<>"掌门" and sf<>"长老" and sf<>"元老" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是掌门或者长老元老，不能发追杀令');}</script>"
	Response.End
end if
if rs("银两")< 100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	ling="[系统]对[##]说：你的银两不够，发一次需要100000银两"
	exit function
end if
ling="<table width='75%' border='0' cellspacing='0' cellpadding='0' align='center'><tr><td height='15' bgcolor='#FF3399'><div align='center'><font color='#FFFFFF' size='4'><font color=yellow face='Wingdings'>[</font><font color='yellow'><b>江 湖 门 派 追 杀 令</b></font><font color=yellow face='Wingdings'>[</font></font></div></td></tr><tr><td bgcolor='#CCCCCC'><div align='center'><font color='#000000' size='2'>"&menpai&sf&aqjh_name&"向本门派弟子下达追杀令:<br>"&fn1&"</font></div></td></tr></table>"
conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>