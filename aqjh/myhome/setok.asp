<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
peiou=clng(trim(Request("peiou")))
guan=clng(trim(Request("guan")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ��ż FROM �û� WHERE ����='" & aqjh_name &"'",conn,1,1
mywife=rs("��ż")
rs.close
if peiou<>0 then
	rs.open "SELECT ���� from [�û�] WHERE ����='" & mywife & "'",conn,1,1
	if rs.Eof or rs.Bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���Ҳ��������ż����,û�а취��������!');javascript:history.go(-1);</script>"
		response.end
	end if
	mywifezh=rs("����")
	rs.close
else
	mywifezh="��"
end if
rs.open "SELECT h_ӵ����2,h_�ι� FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,3
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û�з��ݻ�������Ȩ���ã�');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
rs("h_ӵ����2")=mywifezh
rs("h_�ι�")=guan
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert('��ʾ�����óɹ�!');location.href = 'set.asp';</script>"
%>