<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'解除监禁♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_jc or sjjh_name<>"回首当年" then
	Response.Write "<script language=JavaScript>{alert('提示：解除监禁只有正站长才可以操作！');}</script>"
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
if Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能对其发言。');parent.f2.document.af.towho.value='大家';parent.f2.document.af.towho.text='大家';parent.m.location.reload();</script>"
	Response.end
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【解除监禁】</font><font color=" & saycolor & ">"+shifang(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'解除监禁
function shifang(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 状态 from 用户 where 姓名='"& fn1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你所输入的用户名不存在！');}</script>"
	Response.End
end if
if rs("状态")<>"监禁" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& fn1 &"]并没有被监禁！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 状态='正常',登录=now(),事件原因='"&sjjh_name&" 解除监禁："&fn1&"' where 姓名='" & fn1 & "'"
shifang= "<font color=red>官府的人看"& fn1 &"有心改过，表现良好,经大家一至同意给他一次改过自新的机会，将其释放！</font>"
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & fn1 & "','解除监禁',now(),'有心改过，原谅他这一次！')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>