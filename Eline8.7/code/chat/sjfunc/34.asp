<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'清除♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=green>【清除好友】</font><font color=" & saycolor & ">"+clspy(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'清除
function clspy(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
pengyou=trim(rs("好友名单"))
fn1=trim(fn1)
if instr(pengyou,trim(fn1))=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["& fn1 &"]并不是你的好友，你无法删除！');}</script>"
	Response.End
end if
towho="|" &trim(fn1)& "|"
pengyou=Replace(pengyou,towho,"")
if pengyou="" then
	pengyou=" "
end if
conn.execute "update 用户 set 好友名单='"& pengyou &"' where 姓名='" & sjjh_name &"'"
Response.Write "<script language=JavaScript>{alert('["& fn1 &"]已经从好友列表中清除!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
clspy="##您的好友:"&fn1&"已经从列表中清除………"
end function
%>
