<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
xlwpname=LCase(trim(Request.QueryString("xlwpname")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM x WHERE a='"&xlwpname&"'",conn,3,3
xlwp=rs("b")
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
rs.open "SELECT w6,w8 FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
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
'Response.Write "<script Language=Javascript>alert('��ʾ��"&ii&duyao&duyao1&"');</script>"
'response.end
zstemp=add(rs("w8"),xlwpname,1)
conn.execute "update �û� set w6='"&duyao&"',w8='"&zstemp&"' where  ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("��[<%=xlwpname%>]������ɣ�")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
