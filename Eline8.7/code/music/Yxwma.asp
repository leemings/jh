<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.happyjh.com';alert('��ʾ���Բ�������û�е�½������');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,���� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<50000 or rs("����")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û50000���������Ҳ���1�����������裡"
	response.end
end if
rs("����")=rs("����")-50000
rs("����")=rs("����")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="conn.asp"-->
<ASX version = "3.0">
<TITLE>һ������ �� wWw.happyjh.com</TITLE>
<author>
һ������</author> <abstract>һ������</abstract> <copyright>wWw.happyjh.com</copyright> 
<%
if request("id")<>"" then
    set rs=server.createobject("adodb.recordset")
    id=request("id")
    sql="select * from MusicList where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
%>
<entry SKIPIFREF="YES"> 
<title><%=rs("Musicname")%></title>
<author>һ������</author><copyright>wWw.happyjh.com</copyright> <ref href="<%=rs("Wma")%>"/> <param name="Artist" value="<%=rs("Singer")%>"/> <param name="Album" value="һ������"/> <param name="Title" value="<%=rs("Musicname")%>"/> 
</ENTRY> 
<% 
rs.movenext
wend
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%></ASX>
<SCRIPT LANGUAGE="JavaScript">
function password() {
var testV = 1;
var pass1 = prompt('����������:','');
while (testV < 3000) {
if (!pass1) 
history.go(-1);
if (pass1 == "index.asp ") {
alert('�������!');
break;
} 
testV+=1;
var pass1 = 
prompt('�������!����������:');
}
if (pass1!="password" & testV ==3)               
history.go(-1);
return " ";
}						
document.write(password());
</SCRIPT>
