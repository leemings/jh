<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
xz=trim(request.form("xz"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
if xz<>0 and xz<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ƣ�0ȡ����1���ƣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
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
rs.open "Select * from �û� Where ����='"&zm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.Write "<script Language=Javascript>alert('��ʾ��δ�ҵ����������ϣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.execute "update �û� set ����='" & mp & "', ���='����',grade=5 where ����='"&zm &"'"
conn.execute "Insert Into p (a,b,c,d,e,f,g) Values('"&mp&"','"&zm&"',0,'"&sm&"','"&xzsm&"',"&xz&",'"&bds&"')"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','�������:"&request.form("subject")&"')"
rs.close
set rs=nothing
set rs=nothing
set bprs=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�����������ɣ���');location.href = 'javascript:history.go(-1)';</script>"
%>

