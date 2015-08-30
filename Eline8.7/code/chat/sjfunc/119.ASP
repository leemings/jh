<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'传送道德♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_so then
	Response.Write "<script language=JavaScript>{alert('提示：传送道德需要["&jhdj_so&"]级才可以操作！');}</script>"
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【传送道德】<font color=" & saycolor & ">"+songfali(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'传送道德
function songfali(fn1,to1)
fn1=int(abs(fn1))
if (fn1<100 or fn1>5000) and sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：传送道德最少100点，最多5000点！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 道德 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("道德")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你有那么多的道德吗？看好再说！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 道德=道德-"& fn1 &" where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 道德=道德+"& fn1 &" where 姓名='" & to1 &"'"
songfali=sjjh_name & "把自己的"& fn1 &"点道德传送给了<font color=red>" & to1 & "</font>，" & to1 & "道德大增！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>



