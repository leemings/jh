<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'册封花魁
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('提示：9级才可以操作！');}</script>"
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
'对暂离开、点哑穴判断
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
says="<font color=green>【种花状元开奖】<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'册封花魁
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>1000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：奖品至少得有10个金币！但不得多于1000个！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,银两,花魁 FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
hk=rs("花魁")
if hk>0 then 
Response.Write "<script language=JavaScript>{alert('你想干什么？作弊呀，对方已经是种花状元了，看我不撤了你的官府！');}</script>"
	Response.End
else
conn.execute "update 用户 set 金币=金币+" & fn1 & " where 姓名='" & to1 &"'"
conn.execute "update 用户 set 花魁=20 where 姓名='" & to1 &"'"
end if
givegold="<bgsound src=wav/cfhk.mp3 loop=5><img src=pic/cfhk.gif>经过评比大家一致认同%%为种花状元由##颁发奖品，奖品为金币["&fn1&"]个，大家送花表示祝贺呀。%%在今后的50次登陆聊天室时将会受到特殊的欢迎，如果你也心动也想象人家%%一样那快用你的勤劳和智慧去种花去吧!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','册封种花状元奖励金币"&fn1&"个')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>