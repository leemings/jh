<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'标题功能♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('提示：您不是管理员，不能使用标题功能！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 内力,银两,会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if sjjh_name<>"回首当年" and (rs("内力")<10000 or rs("银两")<5000000) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：需要内力10000、银两500万才可以使用标题！');}</script>"
	Response.End
end if
if rs("会员等级")<2 and sjjh_name<>"回首当年" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('提示：["&sjjh_name &"]你会员不够2级，不能使用标题功能！');window.close();</script>"
 response.end
end if
conn.execute "update 用户 set 银两=银两-5000000,内力=内力-10000 where 姓名='" & sjjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call dianzan(towho)
if len(says)>40 then says=Left(says,40)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+mid(says,i+1)+"(<font color=blue><b>"&sjjh_name&"</b></font>)</font>"
call chatsay("标题",towhoway,towho,saycolor,addwordcolor,addsays,says)
Application.Lock
Application("sjjh_tltie")=says
Application.UnLock
Response.Write "<script language=JavaScript>{alert('提示：标题完成，消耗内力：10000、银两：500万！');}</script>"
%>