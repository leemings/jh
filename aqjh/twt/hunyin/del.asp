<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
Response.End 
end if
if aqjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('��ʾ������������ʲô������ѽ��');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select a from t where b='"& aqjh_name &"' or c='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������û��������ԣ�');history.go(-1);</script>"
	response.end
end if
conn.execute "delete * from t where  id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��ɾ�����弼ID"& id &"�ɹ�');location.href = 'javascript:history.go(-1)';</script>"
%>