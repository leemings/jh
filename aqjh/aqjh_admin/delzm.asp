<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
s=trim(Request.QueryString("username"))
rs.open "select * from p where a='"&s&"'",conn
zm=trim(rs("b"))
rs.close
rs.open "select * from �û� where grade<>10 and ����='"&zm&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=446"
else
	conn.execute("update �û� set ���='����',grade=1 where ����='"&zm&"'")
end if
conn.execute("update p set b='δ��' where a='"&s&"'")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','ȡ����["&s&"]���ţ�["&zm&"]')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=703"
%>