<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'爱神♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<jhdj_qlcy1s then
	call mess("提示：使用爱神传音需要["&jhdj_qlcy1&"]级才可以！",1)
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_jhdj>30 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+titl1(mid(says,i+1))+"</font>"
call chatsay("爱神",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl1(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,名单头像 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
tx=rs("名单头像")
if rs("法力")<500  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：需要法力500好好练几天吧！",1)
end if
conn.execute "update 用户 set 法力=法力-500 where 姓名='" & sjjh_name &"'"
titl1="<bgsound src=wav/folder.WAV loop=1><marquee direction=right><font color=#FF0000>∑☆≡<img src=../ico/"& tx &"-2.gif width=32 height=32>[爱神之箭]" & fn1 & "  (##)" & "</font><font color=#FF0000>≡☆→</font></B></font></MARQUEE>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>