<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../config.asp"-->
<%'扫黄行动
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<25 then
	Response.Write "<script language=JavaScript>{alert('提示：你的等级不够25级,也敢扫黄！小心把你也卖了!');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【扫黄行动】</font><font color=" & saycolor & ">"+saohuang(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function saohuang(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
if to1="大家" or to1=aqjh_name or to1=Application("aqjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('扫黄对象有错，请看仔细了！！');}</script>"
	Response.End
exit function
end if
if rs("性别")<>"男" then
Response.Write "<script language=JavaScript>{alert('扫黄行动暂由女侠执行，男的闪一边！');}</script>"
Response.End
end if
%>
<!--#include file="data.asp"-->
<%
sql="select 姓名 FROM 名妓 WHERE 介绍人='"&to1&"'"
set rs1=connt.execute(sql)
if rs1.eof or rs1.bof then 
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
Response.Write "<script language=javascript>{alert('提示:此人没有从事过拐卖妇女！');}</script>" 
Response.End 
end if 
conn.execute "update 用户 set 银两=银两-5000000 where 姓名='"&to1&"'"
conn.execute "update 用户 set 银两=银两+1000000 where 姓名='"&aqjh_name&"'"
connt.execute "delete * from 名妓 where 介绍人='"&to1&"'"
saohuang=aqjh_name&"和一帮姐妹们带着大刀闯入怡红院把["&to1&"]围起来狠狠的揍了一顿，救出了被["&to1&"]拐卖进去的少女，["&to1&"]被抢掉500万银两，["&aqjh_name&"]得到100万."
rs.close
set rs=nothing
conn.close
set conn=nothing
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
end function 
%>