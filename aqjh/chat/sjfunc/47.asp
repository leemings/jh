<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'求婚
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_jhdj<jhdj_hy then
	Response.Write "<script language=JavaScript>{alert('提示：在线求婚需要["&jhdj_hy&"]级才可以操作！');}</script>"
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
says="<font color=green>【求婚】<font color=" & saycolor & ">"+qiuhun(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'求婚
function qiuhun(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,性别,配偶,道德,魅力,银两 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
sex=rs("性别")
if rs("配偶")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：江湖是一夫一妻制，你想作什么！');}</script>"
	Response.End
end if
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：道德与魅力不够300，人家看不上你的！');}</script>"
	Response.End
end if
if rs("银两")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：没有5万块钱，不能登记结婚的！');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,性别,配偶,等级 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
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
if rs("配偶")<>"无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想作什么，第三者插足呀！');}</script>"
	Response.End
end if
if rs("等级")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：取一个江湖5级的人，没面子吧！');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 银两=银两-50000 where 姓名='" & aqjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
qiuhun="[##]向{%%}求婚：<img src='img/29.gif'>"&fn1&"<input type=button value='我愿意' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('jiehun.asp?name=" & aqjh_name &"&yn=1&to1="&to1&"','d') name=tongyi"&regjm&"><input type=button value='不愿意' onClick=javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('jiehun.asp?name=" & aqjh_name &"&yn=0&to1="&to1&"','d') name=tongyi"&regjm+1&">"
end function
%>
