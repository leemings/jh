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
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))

id=trim(request("id"))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='"&aqjh_name&"'",conn
if rs("ת��")<5 or rs("����")<3000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת��������Σ�����µ���3000�����ܽ�����!');location.href = 'javascript:history.go(-1)';}</script>"
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
zid=rs("id")
 rs.close
select case id
case "�������"

rs.open "select * from [���] where �鳤="&zid&"",conn,1,1
if not(rs.eof or rs.bof) then
id=rs("id")
zuname=rs("����")
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ��������ˣ���id��"&id&"��,������"&zuname&"����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
    conn.execute "insert into ��� (����,�鳤,ʱ��,��Ա,��Ա��) values ('"&aqjh_name&"','"&zid&"','"&now()&"','"&aqjh_name&"',1)"
   	conn.execute "update �û� set ����='"&aqjh_name&"' where ����='" & aqjh_name &"'"

   	says="<font color=red>�������Ϣ��"&aqjh_name&"��������,������"&aqjh_name&"�����Ͽ���ϵ��һ��μ�ɱ���ж��ɡ�</font>"
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ��������ˣ�������"&aqjh_name&"����');location.href = 'index.asp';}</script>"

   end if
 rs.close
case "ɾ�����"
rs.open "select * from [���] where �鳤="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�н����飬��Ҫ�Ҵ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
   			        conn.execute "update �û� set ����='��' where ����='" &aqjh_name&"'"
   			        sqlstr="delete * from ��� where �鳤="&zid&""
conn.Execute sqlstr

   			        
   			        
   			        
   			says="<font color=red>��ɾ����ӡ�"&aqjh_name&"ɾ������,��"&aqjh_name&"����������Ա��ɢ���ұ�����ɡ�</font>"

	Response.Write "<script language=JavaScript>{alert('��ʾ��ɾ����ɹ���������Ա��ɢ��');location.href = 'javascript:history.go(-1)';}</script>"


end if

 rs.close


case "�����Ա"
to1=trim(request("to1"))
rs.open "select ���,���� FROM �û� WHERE ����='"&to1&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"��Ϣ������ȷ���Ƿ��д���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("���")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"û�д����״̬����ֹ���飡');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("����")<>"" and rs("����")<>"��" then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"�Ѿ������ˣ���Ҫ�ظ����룬�����������飬���������飡');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
rs.close
if  Instr(useronlinename,""&to1&"")=0 then
Response.Write "<script Language=Javascript>alert('��" & to1 & "�������������У����ܼ��顣');parent.f2.document.af.towho.options[0].value='���';parent.f2.document.af.towho.options[0].text='���';location.href = 'javascript:history.go(-1)';</script>"
Response.end
end if

rs.open "select * from [���] where �鳤="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�н����飬��Ҫ�Ҵ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
badword=rs("��Ա")
	if InStr(badword,""&to1&"")<>0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"�Ѿ����������˰���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
	if rs("��Ա��")>=Application("aqjh_zudui") then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"���ܼ���,���������Ѿ��ﵽϵͳ���Ƶ�"&Application("aqjh_zudui")&"�ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if



badword=badword&","&to1

 conn.execute "update �û� set ����='"&aqjh_name&"' where ����='" & to1 &"'"
   			       
 sqlstr="update ��� set ��Ա='"&badword&"',��Ա��=��Ա��+1 where ����='" &aqjh_name&"'"
conn.Execute sqlstr

   			says="<font color=#ff0000>�������Ϣ��</font> <font color=blue>�����Ա"&aqjh_name&"�ѣ�"&to1&"���������Լ��飬���Թ�ͬ������</font>"

end if 
rs.close

	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"��ӳɹ���');location.href = 'javascript:history.go(-1)';}</script>"

case "ɾ����Ա"

to1=trim(request("to1"))
if to1=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���鳤������ɾ���Լ�����ɾ�������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

rs.open "select ���,���� FROM �û� WHERE ����='"&to1&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"��Ϣ������ȷ���Ƿ��д���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("���")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"û�д����״̬����ֹ���飡');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("����")<>aqjh_name then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"���������Ա��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
rs.close

rs.open "select * from [���] where �鳤="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�н����飬��Ҫ�Ҵ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
zuname=rs("����")

badword=rs("��Ա")
	if InStr(badword,to1)=0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"�����������˰���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

	badword=Replace(badword,","&to1&"","")

	
	
	
			conn.Execute "update �û� set ����='��' where ����='"&to1&"'"
		conn.Execute "update ��� set ��Ա='"&badword&"',��Ա��=��Ա��-1 where ����='"&zuname&"'"
			conn.Execute "update �û� set ����=����-100 where ����='"&to1&"'"
    			says="<font color=#ff0000>�������Ϣ��</font> <font color=blue>�߳���Ա"&aqjh_name&"����Ա��"&to1&"���߳����飬��"&to1&"���о���û�����ӣ�������ʧ100��</font>"

	end if
	
		Response.Write "<script language=JavaScript>{alert('��ʾ��"&to1&"�߳��ɹ���');location.href = 'javascript:history.go(-1)';}</script>"


rs.close
case else
 Response.Write "<script language=JavaScript>{alert('��ʾ�������д������ȷ����ʹ�ý������ܣ�');location.href = 'javascript:history.go(-1)';}</script>"
 Response.End

end select
					set rs=nothing
			conn.close
			set conn=nothing
		

call showchat(says)
%>

</body>
</html>
