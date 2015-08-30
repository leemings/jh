<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'情人♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_qr then
	Response.Write "<script language=JavaScript>{alert('提示：情人功能需要["&jhdj_qr&"]级才可以操作！');}</script>"
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
says="<font color=green>【情人】<font color=" & saycolor & ">"+qingren(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'情人
function qingren(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 门派,会员金卡,性别,配偶,情人,道德,魅力,银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
sex=rs("性别")
if rs("情人")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您现在有情人了，你要先分手…');}</script>"
	Response.End
end if
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('道德与魅力不够300，人家看不上你的！');}</script>"
	Response.End
end if
if rs("会员金卡")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：要想找一个情人需要2元会员金卡！');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,情人,性别,配偶,等级 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
	Response.End
end if
if rs("性别")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：完了吧，江湖不许同性恋的！');}</script>"
	Response.End
end if
if rs("情人")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：人家有情人了…这样子不好吧…');}</script>"
	Response.End
end if
if rs("等级")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：找一个5及的当情人，没面子！');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 银两=银两-50000 where 姓名='" & sjjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.unLock
randomize()
regjm=int(rnd*3348998)
qingren="<bgsound src=wav/mg.WAV loop=1>[##]向{%%}提出作我情人的请求：<img src='img/ornament02.gif'>"&fn1&"<input  type=button value='接受' onClick="&chr(34)&"javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('qingren.asp?name=" & sjjh_name &"&yn=1&to1="&to1&"','d')"&chr(34)&" name=tongyi"&regjm&"><input type=button value='反对' onClick="&chr(34)&"javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('qingren.asp?name=" & sjjh_name &"&yn=0&to1="&to1&"','d')"&chr(34)&" name=tongyi"&regjm+1&">"
end function
%>
