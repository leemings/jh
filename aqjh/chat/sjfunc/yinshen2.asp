<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
uname=request.Form("nname")
if uname="" then
	response.Write "<script language=javascript>alert('������Ҫ�������û�!');window.location.href='javascript:history.go(-1)';</script>"
	response.End()
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&uname&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	'Response.Redirect "chaterr.asp?id=001" 
	response.Write "<script language=javascript>alert('���û������߻�û�и��û�!');window.location.href='javascript:history.go(-1)';</script>"
	response.End()
end if 
towho="���"
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
says="<font color=green>����ץ������<font color=" & saycolor & ">"+yinshen(mid(says,i+1))+"</font>"
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function yinshen(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ת��,����ʱ�� from �û� where ����='" & uname &"'",conn
zhuans=rs("ת��")
conn.execute "update �û� set ����ʱ��=now() where ����='" & uname &"'"

rs.close
set rs=nothing
conn.close
set conn=nothing

optype=request.Form("op")
if optype="��������" then
if Instr(Application("hidden_man"),""&uname&"")=0  then 
if Application("hidden_man")=""  then
 Application("hidden_man")="ת��������Ա�б�" 
else
Application("hidden_man")=Application("hidden_man")&","&uname
end if
Response.Write "<script language=JavaScript>alert('�㽫"&uname&"�������������У�');parent.top.location.reload();</script>"
Response.End
else
Response.Write "<script language=JavaScript>alert('"&uname&"��û�����߰���');window.location.href='javascript:history.go(-1)';</script>"
Response.End
end if
elseif optype="�ϳ���ʾ��" then
if Instr(Application("hidden_man"),""&uname&"")<>0  then 
hiddenman=Application("hidden_man")
Application("hidden_man")=Replace(hiddenman,","&uname&"","")
Response.Write "<script language=JavaScript>alert('�㽫"&uname&"������ʾ���ˣ�');parent.top.location.reload();</script>"
Response.End
else
Response.Write "<script language=JavaScript>alert('"&uname&"��û��������');window.location.href='javascript:history.go(-1)';</script>"
Response.End
end if 
end if
end function
%>
</body>
</html>
