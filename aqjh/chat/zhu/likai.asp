<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<!--#include file="../../showchat.asp"-->

<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"

id=trim(request("id"))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,����,id FROM �û� WHERE ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ϣ������ȷ���Ƿ��½����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("���")=false then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������״̬!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("���")=""  or rs("���")="��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û����������,״̬����!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
zuname=rs("����")
zid=rs("id")
rs.close
 rs.open "select * FROM ��� WHERE ����='"&zuname&"'",conn
if rs.eof or rs.bof then

	conn.Execute "update �û� set ����='��' where ����='"&aqjh_name&"'"

        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ϣ����������Ϣ��״̬����ʼ��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("�鳤")=zid then 
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���鳤�������߳��Լ�������ɾ�������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
zuzhang=rs("�鳤")

badword=rs("��Ա")
	if InStr(badword,""&aqjh_name&"")=0 then
		conn.Execute "update �û� set ����='��' where ����='"&aqjh_name&"'"

        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������޳�Ա["&aqjh_name&"]��״̬����ʼ��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

	
	badword=Replace(badword,","&aqjh_name&"","")

	
	
	
			conn.Execute "update �û� set ����='��' where ����='"&aqjh_name&"'"
		conn.Execute "update ��� set ��Ա='"&badword&"',��Ա��=��Ա��-1 where ����='"&zuname&"'"

			conn.Execute "update �û� set ����=����-500,allvalue=allvalue-50 where ����='"&zuzhang&"'"


    			says="<font color=#ff0000>�������Ϣ��</font> <font color=blue>�뿪��"&aqjh_name&"�뿪�飨"&zuname&"�����鳤�о���û�����ӣ�������ʧ500���������50�㡣</font>"
rs.close

set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ�������ɹ����Ѿ��뿪���');location.href = 'javascript:history.go(-1)';}</script>"
	

call showchat(says)

%>

</body>
</html>