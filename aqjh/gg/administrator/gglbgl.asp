<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or aqjh_name<>application("aqjh_user") then Response.Redirect "../../error.asp?id=439"
id=request.form("id")
lbmc=request.form("lbmc")
tj=request.form("tj")
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
if tj="�޸�" then
	conn.execute "update l set lb='"&lbmc&"' where id="&id
	conn.execute "update ���� set �������='"&lbmc&"' where �������="&id
elseif tj="ɾ��" then
	conn.execute "delete * from l where id="&id
	conn.execute "update ���� set �������='"&lbmc&"' where �������="&id
elseif tj="����" then
	rs.open "select * from l where id="&id&" or lb='"&lbmc&"'",conn
	if rs.eof or rs.bof then
		conn.execute "insert into l (id,lb) values ("&id&",'"&lbmc&"')"
	else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language=JavaScript>{alert('Ҫ��ӵĹ���ID��"&id&"�����"&lbmc&"�Ѵ��ڣ����������ظ���𣡡�');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
	end if
	rs.close
end if
set rs=nothing
conn.close
set conn=nothing
response.write "<script language=JavaScript>{alert('�������"&tj&"��ϣ���');location.href = 'ggadmin.asp';}</script>"
response.end
%>
