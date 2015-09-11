<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'学成下山
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=green>【学艺有成】</font><font color=" & saycolor & ">"+xuechen(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'学成下山
function xuechen(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 师傅,银两 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("师傅")<>fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：此人不是你师傅或你没有师傅！');}</script>"
	Response.End
end if
if rs("银两")<2000000 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有二百万的离山费！不得私自下山！');}</script>"
	Response.End
end if
conn.execute "update 用户 set 师傅='无' where 姓名='"&aqjh_name&"'"
xuechen="<font color=red>["&aqjh_name&"]经过苦练，在["&fn1&"]的严教下，终于学艺有成，交纳了二百万离山费后，离山了...江湖险恶，以后的日子要看自己了!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>