<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhhy/workauto/mine")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if 
wpsl=int(Session("aqjh_jl")/3)
if wpsl<=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����ֻ�ɵ�0����ô�棿');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,����,����,���� from �û� where ����='"&aqjh_name&"'",conn,2,2
zswp=add(rs("w6"),"��ʯ",wpsl)
mcsl=mywpsl(zswp,"��ʯ")
if rs("����")<wpsl*200000 then
	if rs("����")>1000 and rs("����")>500 then
		aatl=int(rs("����")/2)
		aanl=int(rs("����")/2)
		conn.execute "update �û� set ����="&aatl&",����="&aanl&" where id="&aqjh_id
	else
		conn.execute "update �û� set ����=100,����=100 where id="&aqjh_id
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ǯû������Ҳ��Ҫ���˲ɿ�������������׸ɣ���ȥ��Ǯ������������ǣ����һ����ȥ������һ');</script>"	
else
conn.execute "update �û� set w6='"&zswp&"',����=����-" & wpsl*200000 & " where  ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Session("aqjh_jl")=0
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&"<div align='center'>"&mcsl&"</div>"&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"<div align='center'>0</div>"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"<div align='center'></div>"&chr(34)&"</script>"
Response.Write "<script Language=Javascript>alert('��ʾ���ɵ���["&wpsl&"]��������:["&mcsl&"]����ע�⣺��������������֧��"& wpsl*200000 &" �Ĺ��˹��ʣ����㲻���ò�����ʯ�����Ṥ�˴���ܶ�������������');</script>"
response.end
end if
%>