<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
select case Request.QueryString("rst")
case "1"
response.write "<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>������޸��˷������ã��밴����<a href='room_reset.asp?rst=2'>���������б�</a>��<font color=red>��ע���������������ɾ���˷��䣬���������һ�½������������ֲ���Ҫ�ĺ����</font>"
response.end
case "2"
Application("aqjh_room")=""
Application("aqjh_npc")=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM r",conn
 do while Not rs.Eof
Application("aqjh_room")=Application("aqjh_room")&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&"|"&rs("l")&"|"&rs("m")&";"
Application("aqjh_npc"&rs("a"))=""
	rs.MoveNext
 loop
rs.close
set rs=nothing 
conn.close
set conn=nothing
end select
%>
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>�����������<br><br>������ת�����俴���Ƿ�����