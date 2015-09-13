<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'欢迎新人
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
'if aqjh_grade<9 then
'says=bdsays(says)
'end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)

says="<bgsound src=wav/system.WAV loop=1><font color=green>【欢迎新人】<font color=" & sayscolor & ">"+sjgg(mid(says,i+1))+"</font>"
'Application("aqjh_sjgg")=sjgg(mid(says,i+1)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)



function sjgg(fn1)
sjgg="欢迎新人光临快乐江湖总站，希望这里能成为你温馨的家！~请使用功能一区的♂新人福神♀,有问题请在论坛里发表</a>！<marquee behavior=alternate>" & fn1 & " 颁布：<font color=red>##</font></marquee>"
end function 
%>
