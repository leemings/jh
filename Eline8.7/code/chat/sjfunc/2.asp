<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'点穴♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_dx then
	Response.Write "<script language=JavaScript>{alert('提示：点穴需要管理等级["&grade_dx&"]才可以操作！');}</script>"
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【点穴】</font><font color=" & saycolor & ">"+dian(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dian(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
fn1=Replace(fn1,"=","")
fn1=Replace(fn1,"untion","")
fn1=Replace(fn1,chr(39),"")
rs.open "select grade,姓名 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没有这个人？你是不是看错了！');}</script>"
	Response.End
end if
if rs("grade")<sjjh_grade  then
	conn.execute "update 用户 set 登录=now()+0.02,状态='点穴',事件原因='"&sjjh_name&" 点穴:"&fn1&"' where 姓名='" & to1 &"'"
	dian="##对%%使用了点穴术，%%呆呆地不动了……需28分钟才可以上线……原因："&fn1
	call boot(to1,"点穴，操作者："&sjjh_name&","&fn1)
else
	dian="##对%%使用了点穴术，可是他是官府中高级管理员,你对他没有办法…"
end if
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','点穴','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
