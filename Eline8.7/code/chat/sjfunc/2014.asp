<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'绝命钩♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【绝命钩】<font color=" & saycolor & ">"+xdz(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xdz(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别,w6,等级,体力,内力,职业,法力 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
if sjjh_name=to1 or to1="大家"  then
 xdz=sjjh_name & "你吃饱了撑着了?!不能向大家或自己使用绝命钩呀！"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用绝命钩！！请去职业转换为魔法师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"绝命钩",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你连一个多情环也没有耶.看清楚好吗!');</script>"
	response.end
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用宝物你还得练练呢要求80级以上！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("法力")<10000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,状态,内力,体力 FROM 用户 WHERE 姓名='" & to1 &"'",conn
if rs("等级")>500 then
xdz=to1&" 二眼一瞪，怎么你还想用绝命钩暗杀我吗？？！！"
else
rs.close
rs.open "select 等级,w6,法力,攻击,防御 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
duyao=abate(rs("w6"),"绝命钩",1)
conn.execute "update 用户 set w6='"&duyao&"',法力=法力-10000 where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 体力=体力/10,内力=内力/10  where 姓名='" & to1 &"'"
xdz="<font color=red size=2>" & sjjh_name & "</font>对 " & to1 & " 实施了暗杀行动，好在" & sjjh_name & "功力不高，不然"&to1&"早翘了。("&to1&"体力、内力下降了十分之一)....</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>