<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'逮捕
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_db then
	Response.Write "<script language=JavaScript>{alert('提示：逮捕管理等级["&grade_db&"]才可以操作！');}</script>"
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
says="<font color=green>【刑事拘留】</font><font color=" & saycolor & ">"+daipu(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'逮捕
function daipu(fn1,to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('警告：你不可以对站长操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名,grade,门派 from 用户 where 姓名='" & to1 &"'",conn
if aqjh_grade<=rs("grade") and aqjh_name<>Application("aqjh_user") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：失败，你不能对高级管理操作!');}</script>"
	Response.End
end if
conn.execute "update 用户 set 状态='狱',登录=now()+3,事件原因='"&aqjh_name& " 逮捕：" &fn1&"' where 姓名='" & to1 &"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','逮捕',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
daipu="##:<font color=red><bgsound src=wav/oh_no.wav loop=1>官府的人拿了根铁索套在%%的脖子上,连拉带扯的把%%押到了牢房3天时间！" & fn1 & "</font>"
call boot(to1,"逮捕，操作者："&aqjh_name&","&fn1)
end function
%>