<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if isonline(session("aqjh_name"),session("inroom"))=false then
	Response.Write "<script Language=Javascript>alert('��ʾ��ʹ�ñ����ܱ���Ҫ���������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
money=clng(trim(Request("money")))
money1=clng(trim(Request("money1")))
fs=clng(trim(Request("fs")))
if fs<>1 and fs<>2 then
	Response.Write "<script Language=Javascript>alert('��ʾ��ѡ�������д�������ѡ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ��ż FROM �û� WHERE ����='" & aqjh_name &"'",conn,1,1
mywife=rs("��ż")
rs.close
select case fs
case 1
	rs.open "SELECT ���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,1,3
	if rs.Eof or rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('��ʾ������˺��ڽ����������ڣ�');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	if rs("����")<money then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('��ʾ������������ô���������');location.href = 'javascript:history.go(-1)';</script>"
		Response.End	
	end if
	rs("����")=rs("����")-money
	rs.update
	rs.close
	rs.open "SELECT h_����Ǯ�� FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,3
	rs("h_����Ǯ��")=rs("h_����Ǯ��")+money
	rs.update
	if session("myroom")="h_ӵ����" then
		show=session("aqjh_name") &"��" &now() &"���Լ�Ǯ���д���:"& money &"��"
	else
		show=session("aqjh_name") &"��" &now() &"����żǮ���д���:"& money &"��"
	end if
case 2
	rs.open "SELECT h_����Ǯ�� FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,3
	if rs("h_����Ǯ��")<money then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ǯ����û����ô���Ǯ!');javascript:history.go(-1);</script>"
		response.end
	end if
	rs("h_����Ǯ��")=rs("h_����Ǯ��")-money
	rs.update
	if session("myroom")="h_ӵ����" then
		show=session("aqjh_name") &"��" &now() &"���Լ�Ǯ����ȡ��:"& money &"��!"
	else
		show=session("aqjh_name") &"��" &now() &"����żǮ����ȡ��:"& money &"��!"
	end if
	rs.close
	rs.open "SELECT ���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,1,3
	if rs.Eof or rs.Bof then
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script Language=Javascript>alert('��ʾ������˺��ڽ����������ڣ�');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
	rs("����")=rs("����")+money
	rs.update
end select
conn.execute "insert into m(b,a,c,e,d) values ('" & session("aqjh_name") & "',"& application("sysdate") &",'"& mywife &"','" & show & "','����Ǯ��')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert('"& show &"');location.href = 'money.asp';</script>"
%>