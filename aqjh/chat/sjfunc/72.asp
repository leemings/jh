<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'使用配药
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
'#####房间处理#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
'对暂离开、点哑穴判断
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
says=peiyao(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'使用配药
function peiyao(fn1,to1)
fn1=trim(fn1)
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用配药！');}</script>"
	Response.End
end if
f=Minute(time())
if f<00 or f>10 then
	Response.Write "<script language=JavaScript>{alert('现在不是使用配药时间，使用配药为1-10分钟！');window.close();}</script>"
	Response.End 
end if
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr(";炸药;化尸水;吸血虫;蛊虫;化功散;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己进行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("炸药;化尸水;吸血虫;蛊虫;化功散;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家进行操作！');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
rs.close
rs.open "select grade from 用户 where  姓名='" & aqjh_name &"'",conn,2,2
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
rs.close
fn1=trim(fn1)
rs.open "select w8 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
mycard=abate(rs("w8"),fn1,1)
rs.close
rs.open "select d,e from x where  a='"&fn1&"'",conn,2,2
cz=rs("d")
jjsj=rs("e")
rs.close
select case fn1
case "舍利子"
	conn.Execute ("update 用户 set 杀人数=0 where 姓名='" & aqjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用配药[<b>"&fn1&"</b>]杀人数清0...</font>"
case "炸药","化尸水"
	conn.Execute ("update 用户 set "&cz&"=int("&cz&"/"&jjsj&") where 姓名='" & to1 &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用[<b>"&fn1&"</b>]攻击%%，使得%%的"&cz&"下降1/"&jjsj&"点...</font>"
case "吸血虫","蛊虫","化功散"
	rs.open "select "&cz&" from 用户 where  姓名='"&to1&"'",conn,2,2
	abc=rs(cz)
	rs.close
	conn.Execute ("update 用户 set "&cz&"=int("&cz&"/"&jjsj&") where 姓名='" & to1 &"'")
	conn.Execute ("update 用户 set "&cz&"="&cz&"+int("&abc/jjsj&") where 姓名='" & aqjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用[<b>"&fn1&"</b>]攻击%%，使得%%的"&cz&"下降1/"&jjsj&"点,吸取对方("&cz&")计:+<font color=red>"& int(abc/jjsj) &"</font>点……</font>"
case else
	conn.Execute ("update 用户 set "&cz&"="&cz&"+"&jjsj&" where 姓名='" & aqjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用配药[<b>"&fn1&"</b>]("&cz&")增加：<font color=red>+"&jjsj&"</font>点...</font>"
end select
'删除药品
conn.execute "update 用户 set w8='"&mycard&"' where  姓名='"&aqjh_name&"'"
set rs=nothing	
conn.close
set conn=nothing
end function
%>