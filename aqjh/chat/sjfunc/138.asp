<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'查相同ip
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<grade_ip then
	Response.Write "<script language=JavaScript>{alert('提示：管理需要["&grade_ip&"]级的才可以查看ip的！');}</script>"
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
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=sameip(towho)
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sameip(to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('提示：你想查总站长,找死啊！！！');}</script>"
	Response.End
end if
if to1="心梦缘" then
	Response.Write "<script language=JavaScript>{alert('提示：对方不是江湖中人,不能查！！！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'查用户ip
rs.open "select lastip,regip,grade FROM 用户 WHERE  姓名='" & to1 &"'",conn,2,2
if rs("grade")>aqjh_grade then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('因为他是高级管理员，你无权查看！');}</script>"
	Response.End
end if
ip1=rs("lastip")   '最后ip
ip2=rs("regip")    '注册ip
'查相同注册ip
rs.close
rs.open "select * FROM 用户 WHERE regip='" &ip2&"'",conn,2,2
if rs.recordcount=0 then
   thename2="没有找到任何记录"
else
   thename2="|"
   do while not rs.eof
       name=rs("姓名")
       thename2=thename2&name&"|"
       rs.movenext
   loop
end if
'查相同最后ip
rs.close
rs.open "select * FROM 用户 WHERE lastip='" &ip1&"'",conn,2,2
if rs.recordcount=0 then
   thename1="没有找到任何记录"
else
   thename1="|"
   do while not rs.eof
       name=rs("姓名")
       thename1=thename1&name&"|"
       rs.movenext
   loop
end if
sameip="[查ip]%%查到[<font color=blue><b>"&to1&"</b></font>]的ip地址，和##相同注册IP和最后登录IP有<br>"
sameip=sameip&"[查ip]注册IP:"&ip2&"["&thename2&"]<br>[查ip]最后IP:"&ip1&"["&thename1&"]"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>n
%>