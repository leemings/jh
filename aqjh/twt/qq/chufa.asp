<%@ LANGUAGE=VBScript%>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "qqlist.asp"
userid=request("uid")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,�书 from �û� where ����='"&userid&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=454"
	response.end
end if
if request("id")=1 then
conn.execute "update �û� set ����=����-50000,����=����-500000,�书=�书-5000 where ����='"&userid&"'"
else
conn.execute "update �û� set ����=����-2000000,����=����-20000000,�书=�书-20000 where ����='"&userid&"'"
end if
rs.close
'д��QQ���ݿ�
Set rs1=Server.CreateObject("ADODB.RecordSet")
sql="select * from QQ where ����=0 and oicq="&request("oicq")
rs1.open sql,conn,3,2
if rs1.bof or rs1.bof then response.redirect"qqlist.asp"
rs1("����")=1
if request("id")=1 then
rs1("����˵��")="������50������5���书5ǧ"
mess="������50������5���书5ǧ"
else
rs1("����˵��")="������2000������200���书2��"
mess="������2000������200���书2��"
end if
rs1.update
rs1.close
call showchat("<font color=red>��QQ����������</font>"&userid&"�������󣬱�"&mess)
response.redirect"qqlist.asp?id=3"
%>