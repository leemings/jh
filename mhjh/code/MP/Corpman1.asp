<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then response.redirect "error.asp?id=016"
mypai=session("yx8_mhjh_usercorp")
if mypai="" or mypai="��" or mypai="����" then
response.write "�㻹û�м����κ����ɣ�"
response.end
end if
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM �û� WHERE ����='"&mypai&"'"
Set Rs=conn.Execute(sql)
sql="Select count(*) from �û� where ����='"&mypai&"'"
set rst=conn.execute(sql)
renshu=rst(0)
%>

<head>
<title>�����ֵ�</title>
<link rel="stylesheet" href="../STYLE.CSS">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align="center"><b><font color="#FF3333" size="4" >�����ֵ�</font></b></p>
<table width="99%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center">
<td  width="110" bgcolor="#FFFFFF"> ��<%=mypai%>��[<%=renshu%>]��</td>
<td  width="61" bgcolor="#FFFFFF">�� ��</td> 
<td  width="61" bgcolor="#FFFFFF"> ����</td> 
<td  width="61" bgcolor="#FFFFFF"> �� ��</td> 
<td  width="49" bgcolor="#FFFFFF"> �� ��</td> 
<td  width="56" bgcolor="#FFFFFF">�� ��</td> 
<td  width="61" bgcolor="#FFFFFF">�� ��</td> 
<td  width="49" bgcolor="#FFFFFF">�� ��</td> 
<td  width="63" bgcolor="#FFFFFF">ʦ ��</td> 
</tr> 
<tr> <%do while not rs.bof and not rs.eof%> 
<td width="110"><%=rs("����")%></td> 
<td align="center" width="61"><%=rs("�Ա�")%></td> 
<td align="center" width="61"><%=rs("����")%></td> 
<td align="center" width="61"><%=rs("����")%></td> 
<td align="center" width="49"><%=rs("����")%></td> 
<td align="center" width="56"><%=rs("����")%></td> 
<td align="center" width="61"><%=rs("���")%></td> 
<td align="center" width="49"><%=rs("����")%></td> 
<td align="center" width="63"><%=rs("ʦ��")%></td> 
</tr> 
<% 
rs.movenext 
loop 
rs.Close 
set rs=nothing 
rst.Close 
set rst=nothing 
conn.Close 
set conn=nothing 
%> 
</table> 
