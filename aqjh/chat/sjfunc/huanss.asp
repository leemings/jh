<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'唤回神兽
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
says=Replace(says,"\","我该死")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=#ff5600>【唤回神兽】<font color=" & saycolor & ">"+tiren(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'唤回神兽
function tiren(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
myzs=rs("转生")
zaohuan=rs("召唤兽1")

if myzs<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你真的有神兽了么??');}</script>"
	Response.End
end if
if rs("职业")<>"召唤师" then
	Response.Write "<script language=JavaScript>{alert('没点斤两还想招回神兽！！请去职业转换为召唤师！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if zaohuan="无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你真的有神兽了么??');}</script>"
	Response.End
end if

if to1="大家" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('不要乱来,你的法力还不配支配大家');}</script>"
	Response.End
end if
if to1=rs("姓名") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你肯定有毛病,想自己做神兽');}</script>"
	Response.End
end if
if zaohuan<>to1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('不要乱来,你的法力还不配召唤人类');}</script>"
	Response.End
end if
	tiren="<bgsound src=wav/py.wav loop=1>##看见%%无奈的跟在自己身后,心中烦恼,念起咒语,只见神兽%%仰天一声长啸,凭空消失在空气中,原来是回到神兽世界睡觉去了"
	call boot(to1,"唤回神兽&，操作者："&aqjh_name&","&fn1)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>