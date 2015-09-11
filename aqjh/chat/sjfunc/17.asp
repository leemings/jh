<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'赠送
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_zs then
	Response.Write "<script language=JavaScript>{alert('提示：赠送需要["&jhdj_zs&"]级才可以操作！');}</script>"
	Response.End
end if
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"聊天大厅" then
 Response.Write "<script language=javascript>{alert('提示：赠送物品请去大厅！');}</script>"
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
says="<font color=green>【赠送】<font color=" & saycolor & ">"+zen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'赠送
function zen(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>2 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：类别|物品名|数量 ');}</script>"
	Response.End 
end if
if not isnumeric(zt(2)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
lb=trim(zt(0))
zswupin=trim(zt(1))
wusl=abs(int(clng(zt(2))))
if wusl=0 or wusl>100 and instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于100！');}</script>"
	Response.End 
end if
if lb<>"w1" and lb<>"w2" and lb<>"w3" and lb<>"w4" and lb<>"w5" and lb<>"w6" and lb<>"w7" and lb<>"w8" and lb<>"w10" then
	Response.Write "<script language=JavaScript>{alert('提示：物品类别不正确！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,"&lb&" from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("银两")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送物品需要5000两手续费！');}</script>"
	Response.End
end if
zstemp=abate(rs(lb),zswupin,wusl)
conn.execute "update 用户 set "&lb&"='"&zstemp&"',银两=银两-5000 where  姓名='"&aqjh_name&"'"
rs.close
rs.open "select "&lb&" from 用户 where 姓名='"&to1&"'",conn,2,2
zstemp=add(rs(lb),zswupin,wusl)
conn.execute "update 用户 set "&lb&"='"&zstemp&"' where  姓名='"&to1&"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','操作',now(),'赠送["&zswupin&"]"&wusl&"个')"
zen="##把自己的[<b><font color=blue>" & zswupin & "</font></b>]赠送给了%%["&wusl&"]个，%%很是感谢,此次操作成功!！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
