<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'主题
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_qlcy then
	call mess("提示：使用千里传音需要["&jhdj_qlcy&"]级才可以！",1)
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
act=0
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
says="<font color=" & saycolor & ">"+titl(mid(says,i+1))+"</font>"
call chatsay("千里",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 内力,名单头像 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
tx=rs("名单头像")
if rs("内力")<1000  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：需要内力1000、好好练几天吧！",1)
end if
conn.execute "update 用户 set 内力=内力-1000 where 姓名='" & aqjh_name &"'"
titl="<bgsound src=wav/anqi1.wav loop=2><marquee border='1' onmouseover=this.stop(); onmouseout=this.start();><img src="&tx &">[千里传音]" & fn1 & "  (##)" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>