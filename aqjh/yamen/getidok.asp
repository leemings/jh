<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������֤�벻��ȷ��Ӧ������:"& regjm &"');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
name=request.form("name")
pass=trim(request.form("pass"))
if name="" or pass="" then
	Response.Write "<script Language=Javascript>alert('��ʾ���ǲ����뿪���ˣ��������Ϳ��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select id,���� from �û� where ����='" & name & "' and ����='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�и����������˰���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����벻�԰����᲻��Ǵ��ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=rs("id")
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��������ס�Լ���ID["&id&"]');window.close();</script>"
%>