<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'坐牢♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_zl then
	Response.Write "<script language=JavaScript>{alert('提示：坐牢需要管理等级["&grade_zl&"]才可以操作！');}</script>"
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【坐牢】</font><font color=" & saycolor & ">"+zuolao(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'坐牢
function zuolao(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
fn1=Replace(fn1,"=","")
fn1=Replace(fn1,"untion","")
fn1=Replace(fn1,chr(39),"")
rs.open "select 姓名,grade from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
if sjjh_grade<=rs("grade") then
	Response.Write "<script language=JavaScript>{alert('请不要对高级管理员操作！');}</script>"
	rs.close
	conn.close
	set rs=nothing	
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 状态='牢',登录=now()+1/144,事件原因='"&sjjh_name&" 坐牢:"&fn1&"' where 姓名='" & to1 &"'"
zuolao= "<font color=red><bgsound src=wav/daipu.wav loop=1>官府的人拿了根铁索套在%%的脖子上,一脚把%%踹进了牢房,呆着吧,你需要坐10分钟的牢" & fn1 & "</font>"
call boot(to1,"坐牢，操作者："&sjjh_name&","&fn1)
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','坐牢',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>