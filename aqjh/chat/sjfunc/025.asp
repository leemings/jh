<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'大吼吼
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says="<font color=" & saycolor & ">"+titl3(mid(says,i+1))+"</font>"
call chatsay("大吼吼",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl3(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("知质")<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("提示：需要知质10、好好练几天吧！",1)
end if
conn.execute "update 用户 set 知质=知质-10 where 姓名='" & aqjh_name &"'"
titl3="<STRONG><img src=img/79.gif><font color='#000000' size='4'>【##<bgsound src=wav/dz.wav loop=1>运足真气对%%大声吼叫】" & fn1 & "</STRONG>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>