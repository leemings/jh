<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'斩首
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
useronlinename=Application("aqjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_zs then
	Response.Write "<script language=JavaScript>{alert('提示：暂首需要管理等级["&grade_zs&"]才可以操作！');}</script>"
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<10 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【斩首】<font color=" & saycolor & ">"+zhanshou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'斩首
function zhanshou(fn1,to1)
if to1=Application("aqjh_user") then
   zhanshou= fn1&":你<font color=red>不能对站长%%实行斩立决!</font>"
   exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名,grade from 用户 where 姓名='" & to1 &"'",conn
if aqjh_grade<=rs("grade") and aqjh_name<>Application("aqjh_user") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不可以对高级管理员操作！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 状态='死',allvalue=0,grade=1,门派='游侠',会员等级=0,会员金卡=0,mvalue=0,事件原因='"&aqjh_name&" 斩首"&fn1&"' where 姓名='" & to1 &"'"
zhanshou= "##:<font color=red>官府的人喊到:人犯%%斩立决!</font>"& fn1
call boot(to1,"斩立决，操作者："&aqjh_name&","&fn1)
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','斩首',now(),'" & fn1 & "')"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>