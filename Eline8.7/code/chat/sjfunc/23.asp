<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�Ķ���wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_zs then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Ķ����ߣ���Ҫ["&jhdj_zs&"]������');}</script>"
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
says=xindong(mid(says,i+1))
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�Ķ�
function xindong(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�ȼ� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	Response.Write "<script language=JavaScript>{alert('��ʾ���Ķ�������1000���㲻��ѽ����');}</script>"
end if
	conn.execute "update �û� set ����=����-1000 where ����='" & sjjh_name &"'"
	xindong="<marquee width=100% behavior=alternate scrollamount=10><font color=red><font face=Webdings class=p>&yuml;</font>���Ķ����ߡ�</font>" & fn1 & "(##)" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>