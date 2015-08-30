<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'宝物术♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【魔法师寻宝】<font color=" & saycolor & ">"+bws(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function bws(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 法力,等级,职业,道德,grade,体力,银两,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("职业")<>"魔法师" then
	Response.Write "<script language=JavaScript>{alert('你只是普通子民呢，不能使用宝物术！！请去职业转换为魔法师！');}</script>"
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

dim js(10)
js(0) ="摇钱树"
js(1) ="多情环"
js(2) ="魔法石"
js(3) ="帅哥令"
js(4) ="美人令"
js(5) ="红玛瑙"
js(6) ="绿玛瑙"
js(7) ="蓝玛瑙"
js(8) ="孔雀翎"
js(9)="绝命钩"
randomize()
myxy = Int(Rnd*10)

rs.close
rs.open "select w6 from 用户 where 姓名='"&sjjh_name&"'",conn
duyao=add(rs("w6"),js(myxy),1)
conn.execute "update 用户 set w6='"&duyao&"',法力=法力-1000,银两=银两-10000,体力=体力-1000 where  姓名='"&sjjh_name&"'"
bws=sjjh_name & "法师折腾了半天后终于得到了"&js(myxy)&"一个魔法宝物" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>


