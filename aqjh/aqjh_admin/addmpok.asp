<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from p Where a='"&mp&"'",conn
if not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ����ڣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "Select * from �û� Where  grade<>10 and ����='"&zm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('��ʾ��δ�ҵ����������ϣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
sqlstr="SELECT * FROM p"
rs.Open sqlstr,conn,1,2
rs.AddNew
rs("a")=mp
rs("b")=zm
rs("c")=0
rs("d")=sm
rs("e")=xzsm
rs("g")=bds
rs("h")=0
rs.Update
conn.execute "update �û� set ����='" & mp & "', ���='����',grade=5 where ����='"&zm &"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','�������:"&request.form("mp")&"')"
rs.close
set rs=nothing
set rs=nothing
set bprs=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�����������ɣ���');location.href = 'javascript:history.go(-1)';</script>"
%>
