<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'官府奖励
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<8 then
	Response.Write "<script language=JavaScript>{alert('想作什么呀，你的管理等级可不够呀！');}</script>"
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>【官府奖励】<font color=" & saycolor & ">"+jianglila(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'奖励
function jianglila(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,grade,身份 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("门派")<>"官府" then
      Response.Write "<script language=JavaScript>{alert('你不是官府人员，想捣乱吗？');}</script>"
end if
fn1=int(abs(fn1))
if fn1>5000000 or fn1<1000  then
Response.Write "<script language=JavaScript>{alert('官府奖励不可以超过500万少不可少于1000的！');}</script>"
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        Response.End
end if
conn.execute "update 用户 set 银两=银两+" & fn1 & " where 姓名='" & to1 & "'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','奖励','官府奖励白银"&fn1&"两')"
jianglila="因("& to1 &")对江湖有贡献，"&aqjh_name & "把官府设立的奖金发给(" & to1 &")"& fn1 & "两！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
