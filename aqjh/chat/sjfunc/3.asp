<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'解穴
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_jx then
	Response.Write "<script language=JavaScript>{alert('提示：解穴需要管理等级["&grade_jx&"]才可以操作！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【解穴】</font><font color=" & saycolor & ">"+jie(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'解穴
function jie(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 状态,登录 from 用户 where 姓名='" & fn1 & "'",conn
if rs("状态")="点穴" and (not rs.eof) and rs("登录")>now() then
	conn.execute "update 用户 set 登录=now(),状态='正常',事件原因='"&aqjh_name& ":解除你点穴!" &"' where 姓名='" & fn1 & "'"
	jie="##对" & fn1 & "使用了解穴术，" & fn1 & "猛然醒过来了……"
else
	jie="##对" & fn1 & "使用了解穴术，可是" & fn1 & "并没有被点穴……"
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& fn1 &"','解穴','解除点穴')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
