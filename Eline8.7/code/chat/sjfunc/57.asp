<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'送花♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_hua then
	Response.Write "<script language=JavaScript>{alert('提示：送花需要["&jhdj_hua&"]级才可以操作！');}</script>"
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
says="<font color=red>【江湖消息】<font color=" & saycolor & ">"+songhua(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'送花
function songhua(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('提示：操作错误，格式如下：[物品名&数量]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
zswupin=trim(zt(0))
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>1000000000 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于1000000000！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 道德,魅力,w7 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if  mywpsl(rs("w7"),zswupin)<wusl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&zswupin&"]的数量不够("&wusl&")个！');}</script>"
	Response.End
end if
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('道德与魅力不够300，人家不会收你的花！');}</script>"
	Response.End
end if
rs.close
rs.open "select i FROM b WHERE a='" & zswupin &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
	Response.End
end if
sm=rs("i")
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)
songhua="<bgsound src=wav/call.WAV loop=1>[##]向{%%}送上"&wusl&"朵<img src='../hcjs/jhjs/images/"&sm&"'>" & zswupin &" ,也不知道人家愿不愿意接收...<input type=button value='接收' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('xianhua.asp?fromname="& sjjh_name &"&yn=1&toname="&to1&"&huaming="&zswupin&"&wpsl="&wusl&"','d') name=tongyi"&regjm&"><input type=button value='拒收' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('xianhua.asp?fromname="& sjjh_name &"&yn=0&toname="&to1&"&huaming="&zswupin&"&wpsl="&wusl&"','d') name=tongyi"&regjm+1&">"
end function
%>
