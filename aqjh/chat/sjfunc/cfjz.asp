<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'册封君主
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if application("aqjh_user")<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：只有正站长才能权力操作！');}</script>"
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
says="<font color=green>【册封君主】<font color=" & saycolor & ">"+cfjz(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'册封君主
function cfjz(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & to1 &"'",conn,2,2
guojia=rs("国家")
if rs("等级")<10 then
  		Response.Write "<script language=JavaScript>{alert('提示:对方等级太低，不能胜任君主的职位！');}</script>"
		Response.End
end if
if rs("门派")="官府" then
  		Response.Write "<script language=JavaScript>{alert('提示:官府中人不得参与君主竞选！');}</script>"
		Response.End
end if
if Instr(Application("aqjh_admin_send"),"|" & to1 & "|")<>0 then 
  		Response.Write "<script language=JavaScript>{alert('提示:对方是财神啊！');}</script>"
		Response.End
end if
conn.execute "update 用户 set 职位='君主' where 姓名='" & to1 &"'"
cfjz="<bgsound src=wav/cfhk.mp3 loop=5><img src=pic/cfhk.gif>经过站长审核钦批册封%%为[" & guojia &"]的君主由##颁发奖品，大家热烈祝贺呀。%%在今后登陆聊天室时将会受到特殊的欢迎!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','册封君主')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>