<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'自杀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【自杀】<font color=" & saycolor & ">"+zs(towho,mid(says,i+1))+"</font>"
towhoway=0
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function zs(to1,fn1)
if aqjh_grade>11 then
   zs=aqjh_name&"身为官府人员，既然想不开，不知是为情所困还是受够了当官府的艰辛!"
   exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
sql="update 用户 set 道德=道德+1000,状态='死' where 姓名='" & aqjh_name & "'"
conn.execute sql
call boot(aqjh_name,"自杀，操作者："&aqjh_name&","&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'" & aqjh_name & "','自杀','人命')"
zs=aqjh_name&"大声狂吼：我受不了了....狂奔到快乐江湖最高的生死崖奋不顾身跳了下去,惨啊，又是人间悲剧，[因为勘破生死玄关，道德增加1000点。]"
conn.close
set conn=nothing
end function
%>