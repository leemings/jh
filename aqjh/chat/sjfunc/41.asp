<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'查ip
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
says=seeip(towho)
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'查ip
function seeip(to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('提示：你想查总站长,找死啊！！！');}</script>"
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
sip=split(rs("lastip"),".")
sip1=split(rs("regip"),".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'查最后ip的地区
rs.close
rs.open "select c,d FROM z WHERE a<="& num &" and b>="&num,conn,2,2
if rs.eof or rs.bof then
	country="亚洲"
	city="未知"
else
	country=rs("c")
	city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
last="地区:"& country &" 城市:"& city
rs.close
'查注册ip的地区
num=cint(sip1(0))*256*256*256+cint(sip1(1))*256*256+cint(sip1(2))*256+cint(sip1(3))-1
rs.open "select * FROM z WHERE a<="& num &" and b>="&num,conn,2,2
if rs.eof or rs.bof then
	country="亚洲"
	city="未知"
else
	country=rs("c")
	city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
reg="地区:"& country &" 城市:"& city
seeip="[查ip]%%查到[<font color=blue><b>"&to1&"</b></font>]的现在ip地址为:"&ip1&",鉴定为："&last&"  注册ip地址为:"&ip2 &"鉴定为："&reg
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>