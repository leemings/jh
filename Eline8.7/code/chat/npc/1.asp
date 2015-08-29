<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%'传授轻功
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_im then
	Response.Write "<script language=JavaScript>{alert('提示：传授轻功需要["&jhdj_im&"]级才可以操作！');}</script>"
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

'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【传授轻功】<font color=" & saycolor & ">"+impart(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传授轻功
function impart(fn1,to1)
fn1=int(abs(fn1))
if (fn1<100 or fn1>500000) and sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：传授轻功最少100点，最多5万点！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 轻功 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("轻功")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的轻功吗？看好再说！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 轻功=轻功-"& fn1 &" where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 轻功=轻功+"& fn1 &" where 姓名='" & to1 &"'"
impart=sjjh_name & "把自己的"& fn1 &"点轻功传授给了<font color=red>" & to1 & "</font>，" & to1 & "轻功大涨！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
