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
rs.open "select * from 用户 where grade<>10 and 姓名='"&zm&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=446"
else
	conn.execute("update 用户 set 身份='弟子',grade=1 where 姓名='"&zm&"'")
end if
conn.execute("update p set b='未定' where a='"&s&"'")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','取消：["&s&"]掌门：["&zm&"]')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=703"
%>