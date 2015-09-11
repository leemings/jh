<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'通缉犯
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_tjf then
	Response.Write "<script language=JavaScript>{alert('提示：通缉犯需要["&grade_tjf&"]级以上管理员操作！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/解除" then
	says="<font color=green>【解除通缉】</font><font color=" & saycolor & ">"+tongji(mid(says,i+1),mycz)+"</font>"
else
	says="<font color=green>【通缉人犯】</font><font color=" & saycolor & ">"+tongji(mid(says,i+1),mycz)+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'通缉犯
function tongji(fn1,mycz)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名,grade,通缉 from 用户 where 姓名='" & fn1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你所输入的用户名不存在！');}</script>"
	Response.End
end if
if aqjh_grade<=rs("grade") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('不可以对高级管理员操作！');}</script>"
	Response.End
end if
if mycz="/解除" then
	if rs("通缉")<>True then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示:他也没有被通缉,无法操作！');}</script>"
		Response.End
	end if
	conn.execute "update 用户 set 通缉=False,保护=True where 姓名='" & fn1 &"'"
	tongji="##:<font color=red>宣布"&fn1&"已经被官府解除攻击状态，多年的流落生涯今天总算是到头了~~"
else
	if rs("通缉")<>False then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示:他已经被通缉,无法操作！');}</script>"
		Response.End
	end if
	conn.execute "update 用户 set 通缉=True,保护=False where 姓名='" & fn1 &"'"
	tongji="##:<font color=red>宣布"&fn1&"事坏作尽，多行不义，现在已被官府通缉，杀死通缉犯的大侠将得到100万~~"
end if
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & fn1 & "','操作',now(),'被管理员"&mycz&"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>