<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
wqname=LCase(trim(Request.QueryString("wq")))
if wqname="" then
	Response.Write "<script language=javascript>{alert('��ʾ������ȷѡ����Ҫ������������');parent.history.go(-1);}</script>" 
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM x WHERE a='"&wqname&"'",conn,3,3
xlwp=rs("b")
money=rs("g")
tj=rs("h")
if isnull(xlwp) or xlwp="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������������ⲻ�ܲ���\n�ҳ��򿪷�����ϵ!');</script>"
	response.end
end if
xadata=split(xlwp,"|")
xadata1=UBound(xadata)
rs.close
rs.open "select w6,w3,���� from �û� where ����='"&aqjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻�߱���������\n["&replace(tj,chr(39),"")&"]��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("����")<money then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������["&money&"]��!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ʲô��ƷҲû��!');</script>"
	response.end
end if
for ii=0 to xadata1-1
	xadata2=split(xadata(ii),":")
	mysl=clng(xadata2(1))
	myxlwp=trim(xadata2(0))
	duyao=abate(duyao,myxlwp,mysl)
next
zstemp=add(rs("w3"),wqname,1)
conn.execute "update �û� set w6='"&duyao&"',w3='"&zstemp&"' where  ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("��[<%=wqname%>]������ɣ�")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>