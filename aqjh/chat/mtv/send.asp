<%@ LANGUAGE=VBScript.Encode%>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'送歌
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_ti then
	Response.Write "<script language=JavaScript>{alert('送mtv歌曲需要"&jhdj_ti&"级才可以操作！');}</script>"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><img src=img/music.gif><font color=" & saycolor & ">"+mtv(fnn1,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function mtv(fn1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 道德,魅力,银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('道德与魅力不够300，人家不会听你的歌！');}</script>"
	Response.End
end if
Set connt=Server.CreateObject("ADODB.CONNECTION")
connt.open Application("aqjh_usermdb")
sql="select * from swf where id="&fn1
Set Rs=connt.Execute(sql)
randomize()
regjm=int(rnd*3348998)
mtv="["&aqjh_name&"]将一首《"&rs("歌名")&"》(MTV)送给{"&toname&"}欣赏！ <input type=button value='收看' onClick=javascript:shoukan"&regjm&".disabled=1;window.open('mtv/playswf.asp?id="&aqjh_name&"&name="&fn1&"&toname="&toname&"','playswf','scrollbars=no,resizable=no,width=500,height=400') name=shoukan"&regjm&" style='background-color:#86A231;color:FFFFFF;border: 1 double'>"
rs.close
set rs=nothing
Response.Write "<script language=JavaScript>parent.f2.af.mdsx.checked=true;parent.m.location.reload();</Script>"
end function
%>