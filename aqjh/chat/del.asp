<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
cz=LCase(trim(Request.QueryString("cz")))
lb=LCase(trim(Request.QueryString("lb")))
toname=LCase(trim(Request.QueryString("toname")))
name=LCase(trim(Request.QueryString("name")))
if InStr(toname,"'")<>0 or InStr(toname,"=")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End 
end if
if instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������������ʲô������ѽ��');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select "&lb&" from �û� where ����='"&toname&"'",conn,2,2
if cz="1" then
	zstemp=del(rs(lb),name)
else
	zstemp=add(rs(lb),name,10)
end if
conn.execute "update �û� set "&lb&"='"&zstemp&"' where  ����='"&toname&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
if cz="1" then
	Response.Write "<script Language=Javascript>alert('��ʾ��ɾ��["&toname&"]��Ʒ["&name&"]�ɹ�');location.href = 'towupin.asp?toname="&toname&"';</script>"
else
	Response.Write "<script Language=Javascript>alert('��ʾ������["&toname&"]��Ʒ["&name&"]10���ɹ�');location.href = 'towupin.asp?toname="&toname&"';</script>"
end if
%>
