<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'幸运风采♀wWw.happyjh.com♀
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
says="<font color=green>【幸运风采】<font color=" & saycolor & ">"+xingyu(mid(says,i+1))+"</font>"
if instr(says,"恭喜")=0 then
	towhoway=1
	towho=sjjh_name
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'幸运风采
function xingyu(fn1)
if Application("sjjh_xinyu")=0 or isnull(Application("sjjh_xinyu")) then
	randomize()
	rnd1=cint(rnd()*1999)+1
	Application.Lock
	Application("sjjh_xinyu")=rnd1
	Application.UnLock
end if
fn1=int(abs(fn1))
if fn1<0 or fn1>2000 then
	Response.Write "<script language=JavaScript>{alert('输入错误,幸运风采为：1-2000之间的数！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if fn1=Application("sjjh_xinyu") then
	Application.Lock
	Application("sjjh_xinyu")=0
	Application.UnLock
	conn.execute "update 用户 set 银两=银两+5000000 where 姓名='" & sjjh_name &"'"
	conn.close
	set rs=nothing
	set conn=nothing
	xingyu="★E线风采★：恭喜##您中了E线风采福利采票，号码："&fn1&"<img src='img/251.gif'>奖金500万<img src='img/251.gif'>，大家表示恭喜！"
else
	rs.open "select 银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
	if rs("银两")<1000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('没有钱了，没办法了！别买了！');}</script>"
		Response.End
	end if
	conn.execute "update 用户 set 银两=银两-2000 where 姓名='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	xingyu="★E线风采★：##购买了采票，为E线福利事业作贡献！号码："&fn1&"没有中奖,还有下一次，还有机会！"
end if
end function
%>
