<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=trim(Request.form("name"))
money=abs(request.form("money"))
mess=trim(Request.form("mess"))
if name="" or mess="" or money<10000 or name=sjjh_name then
	Response.Write "<script Language=Javascript>alert('��ʾ������Ϊ�գ���ϢΪ�ջ���ַ�̫��!');location.href = 'lihun.asp';</script>"	
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM h WHERE a='" & sjjh_name & "' or b='" & name & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��["&sjjh_name&"���ڽ��������������ˣ���');location.href = 'lihun.asp';</script>"	
	response.end
end if
rs.close
rs.open "SELECT * FROM �û� WHERE ����='"&sjjh_name&"'" ,conn
If DateDiff("d",rs("��������"),date())<30 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����黹����һ���²�����飡');location.href = 'lihun.asp';</script>"	
	response.end
end if
peiou=trim(rs("��ż"))
if peiou<>name then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��"& name &"Ҳ���������ż��!');location.href = 'lihun.asp';</script>"	
	response.end
end if
if rs("����")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Լ�����ô���Ǯ�𣿣�!');location.href = 'lihun.asp';</script>"
	response.end
end if
rs.close
rs.open "SELECT * FROM �û� WHERE trim(����)='" & name & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��["&name&"�ڽ������Ҳ��������ҹ���Ա�������');location.href = 'lihun.asp';</script>"	
	response.end
end if
rs.close
conn.execute "update �û� set ����=����-"&money&" where ����='"&sjjh_name&"'"
conn.Execute "INSERT INTO h (a,b,c,d,e,f) VALUES('"&sjjh_name&"','"&name&"','���',"&money&",now(),'"&mess&"')"
Response.Write "<script Language=Javascript>alert('��ʾ��"& sjjh_name &"������Ǽǳɹ�����');location.href = 'lihun.asp';</script>"	
response.end
%>
