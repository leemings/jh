<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'好友
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
says="<font color=green>【好友】<font color=" & saycolor & ">"+haoyou(towho,mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'好友
function haoyou(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 好友名单 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
pengyou=trim(rs("好友名单"))
fn1=trim(fn1)
select case fn1
case "加入"
if instr(LCase(pengyou),LCase(to1))<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]已经是你的好友了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
pengyou=pengyou& "|" &to1& "|"
conn.execute "update 用户 set 好友名单='"& pengyou &"' where 姓名='" & aqjh_name &"'"
Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]好友已经加入!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End

case "删除"
if instr(pengyou,trim(to1))=0 then
	Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]并不是你的好友，你无法删除！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
towho="|" &trim(to1)& "|"
pengyou=Replace(pengyou,towho,"")
if pengyou="" then
	pengyou=" "
end if
conn.execute "update 用户 set 好友名单='"& pengyou &"' where 姓名='" & aqjh_name &"'"
Response.Write "<script language=JavaScript>{alert('提示：["& to1 &"]好友已经删除!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End

case "查看"
if pengyou="" then
	Response.Write "<script language=JavaScript>{alert('提示：["& aqjh_name &"]你并没有好友！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
haoyou="##好友名单:"&pengyou
end select
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
