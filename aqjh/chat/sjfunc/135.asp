<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'巧嘴娃娃
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_jhdj<18 then
	Response.Write "<script language=JavaScript>{alert('提示：需要等级18级以上！');}</script>"
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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=1
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>【巧嘴娃娃】<font color=" & saycolor & ">"+qzww(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function qzww(fn1,to1)
fn1=left(trim(fn1),6)
Audibles_ID=left(fn1,2)
if right(fn1,4)<>".swf" or not IsNumeric(Audibles_ID) then
	Response.Write "<script language=JavaScript>{alert('提示：命令出错，请按照格式输入！');}</script>"
	Response.End
end if
if Audibles_ID<=0 or Audibles_ID>60 then
	Response.Write "<script language=JavaScript>{alert('提示：命令出错，请按照格式输入！');}</script>"
	Response.End
end if
aqswf="aq_swf/"&Audibles_ID&".swf"
randomize timer
ki=int(rnd*99)+1
A_id="Audible"&ki&Audibles_ID
swf="<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width=48 height=48 id='"&A_id&"'><param name='movie' value='"&aqswf&"'><param name='quality value='high><param name='wmode' value='transparent'></object><a onmouseover='"&A_id&".Play()'> <font color=red>※</font>鼠标移到这里可以播放^_^</a>"
qzww="##对%%说:<br>"&swf
end function
%>