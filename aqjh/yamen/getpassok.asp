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
userid=request.form("userid")
pass=trim(request.form("pass"))
if name="" or pass="" then
	Response.Write "<script Language=Javascript>alert('��ʾ���������û����ͱ������룡');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select �ڶ����� from �û� where id="&userid&" and ����='" & name & "' and �ڶ�����='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����뱣������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("�ڶ�����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������벻�԰����᲻��Ǵ��ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update �û� set ����='e10adc3949ba59abbe56e057f20f883e' where ����='"&name&"'"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�����ס��������[123456]������ĵ���');window.close();</script>"
%>