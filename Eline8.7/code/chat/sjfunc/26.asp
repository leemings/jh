<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'收徒♀wWw.51eline.com♀
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
'对暂离开、点哑穴判断
call dianzan(towho)
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
says="<font color=green>【收徒】<font color=" & saycolor & ">"+stu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'收徒
function stu(to1)
if trim(Application("sjjh_bais_sf"))<> trim(sjjh_name) then
	Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]也没有想拜你为师！');}</script>"
	Response.End
end if
if trim(Application("sjjh_bais_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]也没有想拜你为师！');}</script>"
	Response.End
end if
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
conn.execute "update 用户 set 银两=银两-50000,师傅='"& sjjh_name &"',师傅交钱='0' where 姓名='"& Application("sjjh_bais_td") &"'"
stu=Application("sjjh_bais_td") & "向##交纳了5万块拜师费，又是点头又是哈腰的，终于求得##收自己为徒,只要师傅在，武功大进的！"
Application.Lock
Application("sjjh_bais_sf")=""
Application("sjjh_bais_td")=""
Application.UnLock
conn.close
set conn=nothing
end function
%>