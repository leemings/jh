<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'竞标商品
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
'#####房间处理#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=jbsp(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'竞标商品
function jbsp(fn1)
fn1=trim(fn1)
zt=split(fn1,"|")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('提示：格式为：竞标物品名|价钱 ');}</script>"
	Response.End 
end if
'取出竞标物品名
wpname=trim(zt(0))
money=abs(int(clng(zt(1))))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())<>1 or (Hour(time())<20 or Hour(time())>=21) then
'if Weekday(date())=7 then
	rs.open "select r FROM b WHERE a='竞标'",conn,2,2
	'判断是否完成竞标，如完成确定竞标。
	if rs("r")=True then
		conn.execute "update b set r=False where a='竞标'"
		rs.close
		rs.open "select a,n,q,o FROM b WHERE n<>'' and q>=10000 and r=True",conn,2,2	
		do while not rs.bof and not rs.eof
			jbname=jbname&"<font color=blue>"&rs("n")&"</font>经营[<font color=blue><b>"&rs("a")&"</b></font>]投资价:<font color=red>"&rs("q")&"两</font>  "
			conn.execute "update 用户 set 存款=存款-"&rs("q")&" where 姓名='"&rs("n")&"'"
			rs.movenext
		loop
		'标记所有竞标商品为已竞标(如果有未竞标的则不标记)
		conn.execute "update b set r=False,p='经营者'&n&'说:欢迎大家选购我的商品，谢谢支持!' where n<>'' and q>=10000"
		conn.execute "update b set r=False"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		jbsp="<font color=green><b>【竞标商品结果】</b><font color=" & saycolor & ">"&jbname&"商品竞标完成，各位获得经营权的玩家可以管理了……</font>"
		exit function
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：竞标商品只可以在周日20-21点进行！');}</script>"
	Response.End
end if
'对已经竞标的商品结算
rs.open "select r FROM b WHERE a='竞标'",conn,2,2
'判断是否完成竞标，如完成确定竞标。
if rs("r")=False then
	conn.execute "update b set r=True where a='竞标'"
	rs.close
	rs.open "select a,n,q,o FROM b WHERE n<>'' and o>=1 and r=False",conn,2,2	
	do while not rs.bof and not rs.eof
		jbname=jbname&"经营[<font color=blue><b>"&rs("a")&"</b></font>]的<font color=blue>"&rs("n")&"</font>收入:<font color=red>"&rs("o")&"两</font>  "
		conn.execute "update 用户 set 存款=存款+"&rs("o")&" where 姓名='"&rs("n")&"'"
		rs.movenext
	loop
	'标记所有竞标商品为已竞标(如果有未竞标的则不标记)
	conn.execute "update b set r=True,n='无',q=0,o=0,p='公告无!' where b='鲜花' or b='药品'"
	conn.execute "update b set r=True"
	rs.close
	set rs=nothing
	conn.close 
	set conn=nothing
	jbsp="<font color=green><b>【竞标商品分红】</b><font color=" & saycolor & ">"&jbname&"</font>现在江湖各位玩家可以竞标自己喜欢的商品了…"
	exit function
end if
'进行竞标系统
rs.close
rs.open "select a,n,q,o FROM b WHERE n='"&aqjh_name&"'",conn,2,2	
trmoney=0
do while not rs.bof and not rs.eof
	trmoney=trmoney+rs("q")
	rs.movenext
loop
rs.close
rs.open "select 存款 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,2,2
if rs("存款")<(money+trmoney) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if trmoney=0 then
		Response.Write "<script language=JavaScript>{alert('提示：你的存款不够["&money&"]两');}</script>"
	else
		Response.Write "<script language=JavaScript>{alert('提示：你已经有投标了现在的钱加上你已经投标的，存款不够["&money&"]两');}</script>"
	end if
	Response.End
end if
rs.close
rs.open "select a,n,q FROM b WHERE a='"&wpname&"' and (b='药品' or b='鲜花') ",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：物品不存或数据出错!');}</script>"
	Response.End
end if
if aqjh_name=rs("n") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：这个现在是你竞标的！!');}</script>"
	Response.End
end if
if money<rs("q") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：竞标价格不够!');}</script>"
	Response.End
end if
conn.execute "update b set n='"&aqjh_name&"',q="&money&",o=0 where a='"&wpname&"'"
jbsp="<font color=green>【竞标商品】<font color=" & saycolor & ">##竞标商品[<b>"&wpname&"</b>]的经营权，此次出价:+<font color=red>"&money&"</font>希望可以获得经营权……</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
