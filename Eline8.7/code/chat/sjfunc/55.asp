<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'单挑♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
if sjjh_jhdj<jhdj_dt then
	Response.Write "<script language=JavaScript>{alert('提示：需要["&jhdj_dt&"]级以上的大侠才可以单挑!');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【单挑武斗】<font color=" & saycolor & ">"+dan(mid(says,i+1),towho)+"</font>"
call chatsay("单挑",towhoway,towho,saycolor,addwordcolor,addsays,says)
'单挑
function dan(fn1,towho)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级 FROM 用户 WHERE 姓名='"&towho&"'",conn,2,2
if rs("等级")<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方等级不够20级不能操作！');}</script>"
	Response.End
end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Application.Lock
Application("sjjh_dantiao")=towho
Application.unLock
dan="[##]向{<font color=blue>%%</font>}提出单挑…原因："&fn1
end function
%>
