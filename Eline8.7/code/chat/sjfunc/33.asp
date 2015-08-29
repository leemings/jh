<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'转账♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_zz then
	Response.Write "<script language=JavaScript>{alert('提示：转账需要["&jhdj_zz&"]级才可以操作！');}</script>"
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
says="<font color=green>【转账】<font color=" & saycolor & ">"+zzyin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'转账
function zzyin(fn1,to1)
fn1=abs(fn1)
if fn1<10000 or fn1>5000000 then
	Response.Write "<script language=JavaScript>{alert('转账最少1万，最多500万！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 存款 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("存款")<fn1 then
	Response.Write "<script language=JavaScript>{alert('你有那么多的存款吗？！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用户 set 存款=存款-"& fn1 &" where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 存款=存款+"& int(fn1*0.95) &" where 姓名='" & to1 &"'"
zzyin="<bgsound src=wav/thanks.WAV loop=1>##把自己江湖的存款:"& fn1 &"两，转账到%%的银行名下，%%向官府交手续费5%此次操作成功！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>