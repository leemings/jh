<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'踢人
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<grade_tr then
	Response.Write "<script language=JavaScript>{alert('提示：踢人管理需要["&grade_tr&"]级的才可以的！');}</script>"
	Response.End
end if
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
says="<font color=#ff5600>【踢人】<font color=" & saycolor & ">"+tiren(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'踢人
function tiren(to1,fn1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('警告：你不可以对站长操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
if rs("grade")<=aqjh_grade or aqjh_name=Application("aqjh_user") then
	tiren="<bgsound src=wav/gf.wav loop=1>一阵狂风把%%刮出了聊天室(聊管=##)原因:"&fn1
	call boot(to1,"踢人，操作者："&aqjh_name&","&fn1)
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('因为他是高级管理员，你不能踢他！');}</script>"
	Response.End
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','踢人','"& fn1 &"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>