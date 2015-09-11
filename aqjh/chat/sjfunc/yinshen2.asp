<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'心情
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
uname=request.Form("nname")
if uname="" then
	response.Write "<script language=javascript>alert('请输入要操作的用户!');window.location.href='javascript:history.go(-1)';</script>"
	response.End()
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&uname&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	'Response.Redirect "chaterr.asp?id=001" 
	response.Write "<script language=javascript>alert('该用户不在线或没有该用户!');window.location.href='javascript:history.go(-1)';</script>"
	response.End()
end if 
towho="大家"
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【鬼抓操作】<font color=" & saycolor & ">"+yinshen(mid(says,i+1))+"</font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'隐身
function yinshen(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 转生,操作时间 from 用户 where 姓名='" & uname &"'",conn
zhuans=rs("转生")
conn.execute "update 用户 set 操作时间=now() where 姓名='" & uname &"'"

rs.close
set rs=nothing
conn.close
set conn=nothing

optype=request.Form("op")
if optype="让他隐身" then
if Instr(Application("hidden_man"),""&uname&"")=0  then 
if Application("hidden_man")=""  then
 Application("hidden_man")="转生隐身人员列表" 
else
Application("hidden_man")=Application("hidden_man")&","&uname
end if
Response.Write "<script language=JavaScript>alert('你将"&uname&"拉进了隐身行列！');parent.top.location.reload();</script>"
Response.End
else
Response.Write "<script language=JavaScript>alert('"&uname&"并没有在线阿！');window.location.href='javascript:history.go(-1)';</script>"
Response.End
end if
elseif optype="拖出来示众" then
if Instr(Application("hidden_man"),""&uname&"")<>0  then 
hiddenman=Application("hidden_man")
Application("hidden_man")=Replace(hiddenman,","&uname&"","")
Response.Write "<script language=JavaScript>alert('你将"&uname&"拉出来示众了！');parent.top.location.reload();</script>"
Response.End
else
Response.Write "<script language=JavaScript>alert('"&uname&"并没有隐身阿！');window.location.href='javascript:history.go(-1)';</script>"
Response.End
end if 
end if
end function
%>
</body>
</html>
