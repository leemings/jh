<%@ LANGUAGE=VBScript codepage ="936" %>
<%response.expires=0%><html>
<head>
<title>NPC�б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<script language="JavaScript">
var chatbgcolor = parent.chatbgcolor;
var chatimage = parent.chatimage;
document.write("<body oncontextmenu=self.event.returnValue=false bgcolor=" + chatbgcolor + " background=" + chatimage + " bgproperties=fixed leftMargin=0 marginwidth=0 marginheight=0 topMargin=1>");
</script>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,n����,n����,n���� from [npc] order by n����",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>û�п����еĶ��� </p>"
else
npcsl=rs.recordcount
%>
<table border="1" cellspacing="0" width="100%" cellpadding="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#85C2E0"> 
    <td align="center"><font color="red">����</font></td>
    <td align="center"><font color="red">�ȼ�</font></td>
  </tr>
  <%do while not rs.eof%>
  <tr bgcolor="#ffffff" onmouseover="this.bgColor='#85C2E0';" onmouseout="this.bgColor='#ffffff';"> 
    <td align="center"><a href="#" onclick="javascript:window.open('npc_see.asp?id=<%=rs("id")%>','seenpc','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=610,height=450')" title="���ˣ�<%=rs("n����")%>,����鿴NPC��ϸ״̬"><%=rs("n����")%></a></td>
    <td align="center"><%=int(sqr(int(rs("n����")/50)))%></td>
  </tr>
<%
rs.movenext
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
<tr bgcolor="#85C2E0" onmouseover="this.bgColor='#ffffff';" onmouseout="this.bgColor='#85C2E0';"> 
<td colspan="2" align="center"><a href="#" onclick="javascript:window.open('npc_read.asp','npc','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=610,height=480')">NPC����</a>[��<%=npcsl%>��]</td></tr>
</table>
</body>
</html>