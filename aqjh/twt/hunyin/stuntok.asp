<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
a=trim(lcase(request("a")))
wg=trim(request.form("wg"))
nei=int(abs(request.form("nl")))
if instr(wg,",")<>0then
	response.write "���ѽ���ڿ���������ز����˰ɣ���"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ��ż<>'��' and ��ż<>���� and ����='"&aqjh_name&"'" ,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�㲢û����ż�����ܴ������弼��"
	response.end
end if
ff=rs("��ż")
sex=rs("�Ա�")
rs.close
rs.open "update �û� set ����=����-10000 where ����='"&aqjh_name&"'",conn
if a="m" then
	id=int(abs(request.form("id")))
	conn.Execute "update t set a='" & wg & "', d=" & nei & " where id="&id
end if
if a="n" then
	if sex="��" then
		conn.Execute "insert into t(a,b,c,d) values ('" & wg & "','" & aqjh_name & "','" & ff & "','" & nei & "')"
	else
		conn.Execute "insert into t(a,b,c,d) values ('" & wg & "','" & ff & "','" & aqjh_name & "','" & nei & "')"
	end if
end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "stunt.asp"
%>