<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=trim(request.form("id"))
a=trim(request.form("a"))
c=abs(int(request.form("c")))
d=abs(int(request.form("d")))
f=trim(request.form("f"))
g=trim(request.form("g"))
if (c+d)>2000 then 
Response.Write "<script language=javascript>{alert('��ʾ���������书�ĺͲ����Դ���2000��!');parent.history.go(-1);}</script>" 
Response.End 
end if 
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhmp")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM p where b='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ѽ���ڿ���������ز����˰ɣ���');location.href = 'javascript:history.back()';</script>"
	response.end
end if
rs.close
sqlstr="SELECT * FROM y where id="&id
rs.Open sqlstr,conn,1,2
rs("a")=a
rs("c")=c
rs("d")=d
rs("f")=f
rs("g")=g
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "setwg.asp"
%>