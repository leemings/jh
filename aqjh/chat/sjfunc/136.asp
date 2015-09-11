<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'挑战掌门
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=green>【挑战掌门】</font><font color=" & saycolor & ">"+tzzm(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tzzm(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
fn1=trim(fn1)
mymp=rs("门派")
mydj=rs("等级")
myzt=rs("体力")+rs("内力")+rs("武功")
if mymp="官府" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：别捣乱！官府人员禁止操作！');}</script>"
	Response.End
end if
if rs("身份")="掌门" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你自己本来就是掌门啊！');}</script>"
	Response.End
end if
if mydj<50 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：等级连50级都没有，你佩作掌门吗？');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM 用户 WHERE 姓名='" &fn1&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：你搞什么，此人并不存在！');}</script>"
	Response.End
end if
tomp=rs("门派")
todj=rs("等级")
tozt=rs("体力")+rs("内力")+rs("武功")
if tomp="官府" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：别捣乱！不能对官府人员操作！');}</script>"
	Response.End
end if
if rs("身份")<>"掌门" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：对方不是掌门啊，你看错了吧？');}</script>"
	Response.End
end if
if mymp<>tomp then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：不是同一门派禁止操作！');}</script>"
	Response.End
end if
if mydj<todj then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的等级太低了，战不过对方！');}</script>"
	Response.End
end if
if myzt<tozt then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的状态太低了，补一补再来吧！');}</script>"
	Response.End
end if
tologintime=CDate(rs("lasttime"))
if tologintime>now()-5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('警告：他在5天内有登录，你不能抢掌门位置！');}</script>"
	Response.End
end if
if rs("银两")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没有100万两银子打理是不行的！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 银两=银两-1000000,grade=5,身份='掌门' where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set grade=1,身份='弟子' where 姓名='"&fn1&"'"
Response.Write "<script language=JavaScript>{alert('恭喜你已经从["&fn1&"]手中夺得掌门之位！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
tzzm="##武功高强，抢得掌门["&fn1&"]的宝座！虽然是抢来的，但还得向官府交纳100万的掌门费，唉！官场真是黑暗啊！随后又说：“小的们，以后就跟着我混了，保证你们吃香的喝辣的！”"
end function
%>