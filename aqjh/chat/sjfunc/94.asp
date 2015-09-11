<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'没收法器
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
says="<font color=red>【没收法器】<font color=" & saycolor & ">"+mushou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'没收法器
function mushou(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,w10,等级,法力 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
dj=rs("等级")

fla=rs("法力")
if rs("法力")<500000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的法力不够无法施展呀，至少也得500000点啊！');window.close();}</script>"
	response.end
end if
if dj<300 then
Response.Write "<script language=JavaScript>{alert('此功能需要300级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 等级,w10 FROM 用户 WHERE 姓名='" & to1 &"'" ,conn,2,2
if rs("等级")<=20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('初入江湖已经够惨了就不要在收他法器吧');}</script>"
	Response.End
end if


duyao=rs("w10")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('对方什么法器都没有啊!');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,w10 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("法力")<2000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：赠送物品需要5000两手续费！');}</script>"
	Response.End
end if
dim js(10)
js(0) ="天堂令"
js(1) ="红钻石"
js(2) ="绿钻石"
js(3) ="蓝钻石"
js(4) ="盗取令"
js(5) ="铁笔银勾"
js(6) ="圣火令"
js(7) ="玄冥棒"
js(8) ="七伤拳"
js(9)="九阳神功"
randomize()
myxy = Int(Rnd*10)
rs.close
rs.open "select w10 from 用户 where 姓名='"&aqjh_name&"'",conn
duyao=abate(rs("w10"),js(myxy),1)
conn.execute "update 用户 set w10='"&duyao&"' where  姓名='"&to1&"'"
duyao=add(rs("w10"),js(myxy),1)
conn.execute "update 用户 set w10='"&duyao&"',法力=法力-2000 where  姓名='"&aqjh_name&"'"
mushou=aqjh_name & "一声大喝：" & to1 & "谁叫你在非动武时间乱用法器伤人的，没有王法了吗，惩戒一下收掉你的"&js(myxy)&"一个" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>