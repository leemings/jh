<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'暴威力豆baodou
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
says="<font color=red>【暴威力豆】<font color=" & saycolor & ">"+baodou()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'暴威力豆baodou
function baodou()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 身份,grade,泡豆点数,暴豆时间 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if trim(rs("身份"))="护法" and rs("grade")=3 then
	doudi=500
else
	doudi=1000
end if
if rs("泡豆点数")<doudi then
	Response.Write "<script language=JavaScript>{alert('提示：你哪里有威力豆？1个豆可以在一小内威力涨3倍！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("暴豆时间"),now())
if sj<=60 then
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('提示："& aqjh_name & "威力时间未过，还剩"& ss &"分钟!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 暴豆时间=now(),泡豆点数=泡豆点数-"& doudi &" where 姓名='" & aqjh_name &"'"
baodou="##口中默念[静心咒]，只见1个绿色的小豆子从头上升起、凝聚、越来越大，嘭！的声炸开，体力、精力得到恢复，从现在开始，60分内##的杀力力会比原来大3倍！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>