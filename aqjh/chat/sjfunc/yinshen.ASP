<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'隐身
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
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
says="<font color=green>【隐身】<font color=" & saycolor & ">"+yinshen(mid(says,i+1))+"</font>"
towho="大家"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'心情
function yinshen(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 转生,操作时间 from 用户 where 姓名='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<8 and aqjh_grade<0 then
	s=8-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您要使用隐身功能请等["&s&"]分再操作！');}</script>"
	Response.End
end if
zhuans=rs("转生")
conn.execute "update 用户 set 操作时间=now() where 姓名='" & aqjh_name &"'"

rs.close
set rs=nothing
conn.close
set conn=nothing
if zhuans<10 and aqjh_grade<10  then 
Response.Write "<script language=JavaScript>{alert('要十转世人才可以用这个功能哇！');}</script>"
Response.End
end if
if Instr(Application("hidden_admin"),""& aqjh_name&"")<>0  then 
Response.Write "<script language=JavaScript>{alert('隐身管理员不需要这个功能哇！');}</script>"
Response.End
end if

if Instr(Application("hidden_man"),""& aqjh_name&"")=0  then 

if Application("hidden_man")=""  then
 Application("hidden_man")="转生隐身人员列表" 

else
Application("hidden_man")=Application("hidden_man")&","&aqjh_name
end if
Response.Write "<script language=JavaScript>{alert('你打开了隐身功能！');}</script>"
Response.Write ("<Script Language=JavaScript>parent.top.location.reload();</Script>")

Response.End
else 
hiddenman=Application("hidden_man")
Application("hidden_man")=Replace(hiddenman,","& aqjh_name&"","")
Response.Write "<script language=JavaScript>{alert('你关闭了隐身功能！');}</script>"
Response.Write ("<Script Language=JavaScript>parent.top.location.reload();</Script>")

Response.End
end if 



end function
%>
</body>
</html>
