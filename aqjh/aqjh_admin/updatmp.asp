<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=trim(request.form("id"))
mp=trim(request.form("mp"))
zm=trim(request.form("zm"))
xz=trim(request.form("xz"))
sm=trim(request.form("sm"))
xzsm=trim(request.form("xzsm"))
bds=trim(request.form("bds"))
jj=trim(request.form("jj"))
bds=replace(bds,"'",chr(34)&"&chr(39)&"&chr(34))
if xz<>0 and xz<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ƣ�0ȡ����1���ƣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if mp="" or zm="" or sm="" or xzsm="" or bds="" or jj="" then
	Response.Write "<script Language=Javascript>alert('��ʾ�������пյ����ݴ��ڣ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from p Where id="&id,conn
yzm=rs("a")
rs.close
rs.open "Select * from p Where a='"&mp&"' and id<>"&id,conn
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
	Response.Write "<script Language=Javascript>alert('��ʾ��δ�ҵ�������["&mp&"]���ϣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
'ȡ��ԭ����
conn.execute "update �û� set ����='" & mp & "', ���='����',grade=1 where ����='"& yzm &"'"
'����������
if trim(mp)<>"�ٸ�" then
	conn.execute "update �û� set ����='" & mp & "', ���='����',grade=5 where ����='"&zm &"'"
else
	conn.execute "update �û� set ����='" & mp & "', ���='����',grade=10 where ����='"&zm &"'"
end if
'������������!
conn.execute "update p set a='" & mp & "',b='" &zm& "',d='" &sm& "',e='" &xzsm& "',f=" &xz& ",g='"& bds &"',h="&jj&" where id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=701"
%>