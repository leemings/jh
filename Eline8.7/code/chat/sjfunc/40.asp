<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'心跳♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_xt then
	Response.Write "<script language=JavaScript>{alert('提示：心跳操作需要战斗等级["&jhdj_xt&"]，才可以操作！！');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=xintiao(mid(says,i+1))
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'心跳
function xintiao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 内力 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("内力")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要内力1000才可以操作！你不够呀！！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 内力=内力-1000 where 姓名='" & sjjh_name &"'"
xintiao="<marquee height=50 behavior=alternate loop=200 direction=up onmouseover=this.stop(); onmouseout=this.start();>〖心跳心语〗" & fn1 & "(##)" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>