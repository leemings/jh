<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'生日
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('需要20等级！');}</script>"
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
says="<font color=green>【生日快乐】<font color=" & saycolor & ">"+lenyin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lenyin(fn1,to1)
if to1=aqjh_name or to1="大家" then
   Response.Write "<script language=JavaScript>{alert('生日对谁？');}</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,会员等级 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("银两")<20000 then
  Response.Write "<script language=JavaScript>{alert('你哪里有20000两银子？');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
hy=rs("会员等级")
if hy=0 then
Response.Write "<script language=JavaScript>{alert('你不是等级会员,请进行购买!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
else
conn.execute "update 用户 set 银两=银两-20000 where 姓名='" & aqjh_name & "'" 
conn.execute "update 用户 set 魅力=魅力+100 where 姓名='" & to1 & "'"
lenyin="##对%%边祝福边唱生日歌，Happy birthday tu you！恭祝你生日也快乐，恭祝你生辰快乐，年年有今朝，岁岁有今日……接着，爱情众多朋友一起祝福%%天天快乐，%%感动的和大家一起度过了美好的一天，一起分享了喜讯，这夜……大家狂欢到4点还依依不舍！" & to1 & "开心的吃了<IMG src='img/1-2.gif'>，魅力增加了<font color=red>100</font>，感叹到，真希望此时常久……！！"
end if
conn.close
set conn=nothing
end function
%>