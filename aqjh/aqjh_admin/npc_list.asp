<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html><title>npc�б�</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
</head>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,n����,n�Ա�,nͼƬ,n����,n����,n����,n�书,n����,n����,n��Ʒ,n����,n��Ʒ����,n��Ʒ,n���� from [npc] order by n����",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>û�п����еĶ��� </p>"
else
npcsl=rs.recordcount
%>
<table border="1" width="90%" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
  <tr bgcolor="#336633"> 
    <td colspan="13" align=center><font color="#FFFFFF">NPC �� ��</td>
  </tr>
  <tr bgcolor="#336633"> 
    <td width="5%" align=center><font color="#FFFFFF">ID��</td>
    <td width="10%" align=center><font color="#FFFFFF">����</td>
    <td width="4%" align=center><font color="#FFFFFF">�Ա�</td>
    <td width="5%" align=center><font color="#FFFFFF">����</td>
    <td width="6%" align=center><font color="#FFFFFF">����</td>
    <td width="4%" align=center><font color="#FFFFFF">����</td>
	<td width="4%" align=center><font color="#FFFFFF">�书</td>
    <td width="9%" align=center><font color="#FFFFFF">����</td>
    <td width="7%" align=center><font color="#FFFFFF">����</td>
    <td width="9%" align=center><font color="#FFFFFF">����</td>
    <td width="5%" align=center><font color="#FFFFFF">����</td>
  </tr>
<%do while not rs.eof%>
  <tr> 
    <td width="5%" align=center><%=rs("id")%></td>
    <td width="10%" align=center><%=rs("n����")%></td>
    <td width="4%" align=center><%=rs("n�Ա�")%></td>
    <td width="5%" align=center><%=rs("n����")%></td>
    <td width="6%" align=right><%=rs("n����")%></td>
    <td width="4%" align=right><%=rs("n����")%></td>
    <td width="4%" align=right><%=rs("n�书")%></td>
    <td width="9%" align=right><%=rs("n����")%></td>
    <td width="7%" align=right><%=rs("n����")%></td>
    <td width="9%" align=left><%=rs("n��Ʒ����")%></td>
    <td width="5%" align=center><a href="npc_show.asp?npcid=<%=rs("id")%>">��ϸ</a></td>
  </tr>
<%
rs.movenext
loop
%>
</table>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>