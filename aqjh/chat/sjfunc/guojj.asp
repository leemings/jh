<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'查看本国基金seegjj
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
says="<font color=red>【本国国库】<font color=" & saycolor & ">"+seegjj()+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'查看本国基金seegjj
function seegjj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select 国家基金,国家 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
guojia=rs("国家")
if guojia="无" or guojia="未知" or guojia="无" or guojia="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你看清楚有没有加入国家，还查看国库呢！');}</script>"
	Response.End
end if
myjj=rs("国家基金")
rs.close
rs.open "select jin FROM guo WHERE g='" & guojia & "'",conn,2,2
bgjj=rs("jin")
seejj="##你现的对本国的贡献：<font color=red><b>"&myjj&"</b></font>两,["&guojia&"]的国库为：<font color=red><b>"&bgjj&"两！</b></font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>