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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ����='"&aqjh_name&"'",conn,2,2
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("����")<>"����" then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('�������ţ���Ҫ�Ҵ���');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
mp=rs("����")
rs.close
rs.open "SELECT * FROM y where a='" & a & "'",conn,2,2
if not(rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������ʽ�Ѿ����ڣ�');location.href = 'javascript:history.back()';</script>"
	response.end
end if
rs.close
sqlstr="SELECT * FROM y "
rs.Open sqlstr,conn,1,2
rs.AddNew
rs("a")=a
rs("b")=mp
rs("c")=c
rs("d")=d
rs("e")=0
rs("f")=f
rs("g")=g
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "setwg.asp"
%>