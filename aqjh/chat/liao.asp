<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
nowinroom=session("nowinroom")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���� from �û� where ����='"& aqjh_name &"'",conn
if rs("����")<100 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('û��100��Ǯ�����ǽ��������������ҵ�Ŷ��');window.close();}</script>"
Response.End
end if
if rs("����")<100 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��������50��ô������������..');window.close();}</script>"
Response.End
end if
conn.execute "update �û� set ����=����-100 where ����='"& aqjh_name &"'"
conn.execute "update �û� set ����=����-50 where ����='"& aqjh_name &"'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
says="<font color=green>���������졿<font color=red>["& aqjh_name&"]����100��������50����������"&Application("aqjh_chatroomname")&"���ģ���ҿ��������ҳ����^_^</font>"
dim no,na,pass
no="740196"
na=session("aqjh_name")
pass=123456
if na<>Application("aqjh_user") then showchat (says)
response.write "<html><head><title>��������</title></head><frameset rows='*,0' frameborder='NO' border='0' framespacing='0' cols='*'><frame name='mainFrame' src='http://202.96.140.86/login.php?USER="&na&"&PASS="&pass&"&roomid="&no&"'><frame name='bottomFrame' scrolling='NO' noresize src=''></frameset><noframes><body bgcolor='#FFFFFF' text=#000000></body></noframes></html>"
%>
