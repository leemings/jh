<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'掌门令♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('掌门令需要4级才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=ling(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'掌门令
function ling(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 身份,银两,门派 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,2,2
sf=rs("身份")
menpai=rs("门派")
if sf<>"掌门" and sf<>"长老" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是掌门或者长老，不能发掌门令');}</script>"
	Response.End
end if
if rs("银两")< 3000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	ling="[系统]对[##]说：你的银两不够，发一次掌门令需要3000000银两"
	exit function
end if
ling="<table width='75%' border='0' cellspacing='0' cellpadding='0' align='center'><tr><td height='15' bgcolor='#FF3399'><div align='center'><font color='#FFFFFF' size='4'><font color=yellow face='Wingdings'>[</font><font color='yellow'><b>掌 门 令</b></font><font color=yellow face='Wingdings'>]</font></font></div></td></tr><tr><td bgcolor='#CCCCCC'><div align='center'><font color='#000000' size='2'>"&menpai&sf&sjjh_name&"曰,门下弟子听令:<br>"&fn1&"</font></div></td></tr></table>"
conn.execute "update 用户 set 银两=银两-3000000 where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>