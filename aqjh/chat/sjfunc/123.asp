<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'冷饮
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
says="<font color=green>【冷饮】<font color=" & saycolor & ">"+lenyin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lenyin(fn1,to1)
if to1=aqjh_name or to1="大家" then
   Response.Write "<script language=JavaScript>{alert('冷饮对谁？');}</script>"
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
conn.execute "update 用户 set 体力=体力+200,内力=内力+200 where 姓名='" & to1 & "'" 
lenyin="##对%%说：好兄弟，渴了吧！请吃<IMG src='img/Z461.gif'>（味道好极了！），呆会再聊！" & to1 & "喝了的冷饮体力和内力都增加了200"
end if
conn.close
set conn=nothing
end function
%>