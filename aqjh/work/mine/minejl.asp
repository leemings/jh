<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"work/mine")=0 then 
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
rs.open "select w6 from �û� where ����='"&aqjh_name&"'",conn,2,2
zswp=add(rs("w6"),"��ʯ",wpsl)
mcsl=mywpsl(zswp,"��ʯ")
conn.execute "update �û� set w6='"&zswp&"' where  ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Session("aqjh_jl")=0
Response.Write "<script Language=Javascript>;parent.show.fm1.innerHTML="&chr(34)&mcsl&chr(34)&";parent.show.fm2.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.fm3.innerHTML="&chr(34)&"0"&chr(34)&";parent.show.shiform.fmsl.value=0</script>"
Response.Write "<script Language=Javascript>alert('��ʾ���ɵ���["&wpsl&"]��������:["&mcsl&"]����');</script>"
response.end
%>