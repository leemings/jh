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
name=trim(request("name"))
my=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h where a='" & name & "' and b='"&my&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=481"
	response.end
end if
rs.close
rs.open "select * from �û� where ��ż='��' and ��ż<>���� and ����='"&name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=430"
	response.end
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sex=rs("�Ա�")
rs.close
rs.open "SELECT * FROM �û� where ����='"&aqjh_name&"'",conn
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
if rs("�Ա�")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=428"
	response.end
end if			
if rs("��ż")<>"��" then
	conn.execute "delete * from h where a='" & name & "' or b='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=429"
	response.end
end if
	conn.execute "update �û� set ��ż='" & name & "',��������=date(),������=������+1 where ����='"&aqjh_name&"'"
	conn.execute "update �û� set ��ż='" & my & "',��������=date(),������=������+1 where ����='" & name & "'"
	conn.execute "delete * from h where a='" & name & "' or b='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "yuelao.asp"
%>