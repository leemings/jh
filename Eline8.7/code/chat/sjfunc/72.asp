<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="chatconfig.asp"-->
<%'使用配药♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房间处理#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用配药！');}</script>"
	Response.End
end if
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="高手E线" then
	Response.Write "<script language=JavaScript>{alert('提示：夺宝大赛中不可以使用配药！');}</script>"
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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
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
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if sjjh_name=to1 and instr(";炸药;化尸水;吸血虫;蛊虫;化功散;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己进行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("炸药;化尸水;吸血虫;蛊虫;化功散;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家进行操作！');</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
rs.open "select 会员金卡,w8 from 用户 where  姓名='"&sjjh_name&"'",conn,2,2
mycard=abate(rs("w8"),fn1,1)
rs.close
rs.open "select d,e from x where  a='"&fn1&"'",conn,2,2
cz=rs("d")
jjsj=rs("e")
rs.close
select case fn1
case "舍利子"
	conn.Execute ("update 用户 set 杀人数=0 where 姓名='" & sjjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用配药[<b>"&fn1&"</b>]杀人数清0...</font>"
case "炸药","化尸水"
	conn.Execute ("update 用户 set "&cz&"=int("&cz&"/"&jjsj&") where 姓名='" & to1 &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用[<b>"&fn1&"</b>]攻击%%，使得%%的"&cz&"下降1/"&jjsj&"点...</font>"
case "吸血虫","蛊虫","化功散"
	rs.open "select "&cz&" from 用户 where  姓名='"&to1&"'",conn,2,2
	abc=rs(cz)
	rs.close
	conn.Execute ("update 用户 set "&cz&"=int("&cz&"/"&jjsj&") where 姓名='" & to1 &"'")
	conn.Execute ("update 用户 set "&cz&"="&cz&"+int("&abc/jjsj&") where 姓名='" & sjjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用[<b>"&fn1&"</b>]攻击%%，使得%%的"&cz&"下降1/"&jjsj&"点,吸取对方("&cz&")计:+<font color=red>"& int(abc/jjsj) &"</font>点……</font>"
case else
	conn.Execute ("update 用户 set "&cz&"="&cz&"+"&jjsj&" where 姓名='" & sjjh_name &"'")
	peiyao="<font color=green>【使用配药】<font color=" & saycolor & ">##使用配药[<b>"&fn1&"</b>]("&cz&")增加：<font color=red>+"&jjsj&"</font>点...</font>"
end select

'删除自己卡片，记录
conn.execute "update 用户 set w8='"&mycard&"' where  姓名='"&sjjh_name&"'"
set rs=nothing	
conn.close
set conn=nothing
end function
%>
