<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'宝物术
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"聊天大厅" then
 Response.Write "<script language=javascript>{alert('提示：找宝物请去大厅，这里没有！');}</script>"
 Response.End
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
says="<font color=red>【寻找珍宝】<font color=" & saycolor & ">"+bws(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function bws(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 法力,等级,职业,道德,grade,体力,银两,操作时间,智力 FROM 用户 WHERE 姓名='" & aqjh_name &"'" ,conn,2,2
if DateDiff("s",rs("操作时间"),now())<60 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('为了防止在乱刷，寻找宝物等30秒钟操作!');}</script>"
	Response.End 
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("等级")<80 then
	Response.Write "<script language=JavaScript>{alert('你的等级不够呀！使用宝物术最少需要80级！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("体力")<1000 then
	Response.Write "<script language=JavaScript>{alert('你的体力不够1000了！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("智力")<10 then
	Response.Write "<script language=JavaScript>{alert('你的智力不够10了！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("银两")<10000 then
	Response.Write "<script language=JavaScript>{alert('你的银两不够了要一万两！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("法力")<1000 then
	Response.Write "<script language=JavaScript>{alert('你法力不够了要1000的法力！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

dim js(11)
js(0) ="摇钱树"
js(1) ="多情环"
js(2) ="绝命钩"
js(3) ="帅哥令"
js(4) ="美人令"
js(5) ="红玛瑙"
js(6) ="绿玛瑙"
js(7) ="蓝玛瑙"
js(8) ="孔雀翎"
js(9) ="美人令"
js(10) ="帅哥令"
randomize()
myxy = Int(Rnd*11)
rs.close
rs.open "select w6 from 用户 where 姓名='"&aqjh_name&"'",conn
duyao=add(rs("w6"),js(myxy),1)
conn.execute "update 用户 set w6='"&duyao&"',法力=法力-1000,智力=智力-10,银两=银两-10000,体力=体力-1000,操作时间=now() where  姓名='"&aqjh_name&"'"
bws=aqjh_name & "法师折腾了半天后终于得到了"&js(myxy)&"一个魔法宝物" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>