<!--#include file="sjfunc.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<100 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ100���ſ��Բ�����');}</script>"
	Response.End
end if
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if towho="���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ������ɣ�����ң���');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","��")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><b>�����ס�</b><font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
bl=rs("��ż")
if bl<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������û�и������������ˣ�������ƨѽ���ٺ٣�');}</script>"
	Response.End
end if
if rs("����")<50000 or rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵ��»���������5�򡭡�');}</script>"
	Response.End
end if
if rs("֪��")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����֪������500����');}</script>"
	Response.End
end if
mysex=rs("�Ա�")
rs.close
rs.open "select * FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ǽ������ˣ������Զ���ʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
if rs("�Ա�")=mysex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ε����Է�����ͬһ�Ա��ѵ����ǡ���');}</script>"
	Response.End
end if
if rs("��ż")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��Ѿ��ɼ��ˣ�������������İɡ���');}</script>"
	Response.End
end if
conn.execute "update �û� set ��ż='"&aqjh_name&"' where ����='"&to1&"'"
conn.execute "update �û� set ��ż='"&to1&"',allvalue=allvalue-int(allvalue/100),����=����-50000,����=����-50000,֪��=֪��-500 where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
Application.Lock
Application("aqjh_tw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
tw="<font color=green>��������ʲô����ѽ��[<font color=red><b>"&aqjh_name&"</b></font>]��Ȼ���������۹⣬��[<font color=red><b>"&to1&"</b></font>]ǿ���ؼң�����[<font color=red><b>"&aqjh_name&"</b></font>]����һ��������[<font color=red><b>�ٷ�֮һ</b></font>]��ͬʱҲʧȥ�˲��ٵ�����������ֵ����<font>"
end function
%>