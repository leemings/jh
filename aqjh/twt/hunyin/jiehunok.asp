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
name=trim(Request.form("name"))
mess=trim(Request.form("mess"))
money=Request.form("money")
if not isnumeric(money) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&money&"]�������ID��ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
money=int(money)
if name=aqjh_name or name="" or mess="" or money<=100000 then Response.Redirect "../../error.asp?id=433"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM h WHERE a='"&aqjh_name&"'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=431"
	response.end
end if
rs.close
'У���û� ��������100��Ǯ����1000
rs.open "SELECT * FROM �û� WHERE ����='"&aqjh_name&"'"&" and ��ż='��' and �ȼ�>5 and ����>"&(money+100000),conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=433"
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
sex=trim(rs("�Ա�"))
rs.close
rs.open "SELECT * FROM �û� WHERE �ȼ�>18 and ����='" & name & "' and ��ż='��' and �Ա�<>'"&sex&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=430" 
end if
	conn.execute "update �û� set ����=����-"&(money+100000)&" where ����='"&aqjh_name&"'"
	conn.Execute "INSERT INTO h (a,b,c,d,e,f) VALUES('"&aqjh_name&"','"&name&"','���',"&money&",now(),'"&mess&"')"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../ok.asp?id=600"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>