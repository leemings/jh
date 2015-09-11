<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'使用
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says="<font color=green>【使用药品】<font color=" & saycolor & ">"+eat(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'使用
function eat(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：物品名|数量 ');}</script>"
	Response.End 
end if
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
zswupin=trim(zt(0))
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 and instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于100！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w1 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
eattemp=abate(rs("w1"),zswupin,wusl)
rs.close
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
conn.execute "update 用户 set w1='"&eattemp&"',内力=内力+"&nl&",体力=体力+"&tl&" where  姓名='"&aqjh_name&"'"
eat="##使用了药品[<b>"&zswupin&"</b>]共计："&wusl&"个,自己的体力：<font color=red>+"&tl&"</font>点,内力:<font color=red>+"&nl&"</font>点!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>