<%
if session("aqjh_name")="" then response.redirect"../../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if request("id")="" then
sql="select * from QQ order by �û��� desc"
elseif request("id")=1 then
sql="select * from QQ where ����<>0 and ����=0"
elseif request("id")=2 then
sql="select * from QQ where ����=0 and ����<>0"
else
sql="select * from QQ where ����=0 and ����=0"
end if
rs.open sql,conn,3,2
if rs.eof and rs.bof then str="��ϵͳ��û��ע���û�!"
if str="" then
	rs.PageSize=20 'ÿҳ����
	pages=rs.pagecount
	records=rs.recordcount
	currentpage=request("currentpage")
	if currentpage="" or currentpage<1 then currentpage=1
	currentpage=cint(currentpage)
	if currentpage>pages then currentpage=pages
	rs.absolutepage=currentpage
else
        currentpage=1
	records=0
	pages=1
end if
%>
<html>
<head>
<title>���������б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF">
<p>�� <a href="qqlist.asp">�������б�</a> 
�� <a href="qqlist.asp?id=1">��������</a> 
�� <a href="qqlist.asp?id=2">��������</a> 
�� <a href="qqlist.asp?id=3">��δ��������</a> 
  <a href="send.asp"><b>���뽱��</b></a></p>
<p align="center"><b><font size="3"><%if request("id")="" then%>���뽱����ȫ������<%elseif request("id")=1 then%>��������<%elseif request("id")=2 then%>��������<%else%>δ��������<%end if%></font></b></p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
  <tr bgcolor="#990000"> 
    <td align="center"><font color="#FFFF00">�û���</font></td>
    <td align="center"><font color="#FFFF00">OICQ����</font></td>
    <td align="center"><font color="#FFFF00">��������</font></td>
    <td align="center"><font color="#FFFF00">�Ƿ����</font></td>
  </tr>    <%linenumber=rs.pagesize%> <%do while (not rs.eof) and (line<linenumber)%> 
  <tr bgcolor="#FFFFFF"> 
    <td><%=rs("�û���")%></td>
    <td><%=rs("oicq")%></td>
    <td align="center"><%=rs("����")%></td>
    <td><%if rs("����")=0 and rs("����")=0 then%>���ڴ�����<%else%><%=rs("����˵��")%> <%=rs("����˵��")%><%end if%> 
      <%if session("aqjh_grade")=10 then%>��<%if rs("����")=0 then%><a href="jiangli.asp?uid=<%=rs("�û���")%>&oicq=<%=rs("oicq")%>"> 
      ����</a><%end if%> <a href="chufa.asp?uid=<%=rs("�û���")%>&oicq=<%=rs("oicq")%>">����</a>1 
      <a href="chufa.asp?id=1&uid=<%=rs("�û���")%>&oicq=<%=rs("oicq")%>">����2</a>����<a href="delqq.asp?uid=<%=rs("�û���")%>&oicq=<%=rs("oicq")%>">ɾ��</a>��<%end if%></td>
  </tr><%rs.movenext%> <%line=line+1%> <%loop%> <%rs.close%> 
</table>
<br>
&nbsp;�Ѽ������ѹ�<font color="blue"><%=Records%></font>λ&nbsp; ��<font color="blue"><%=Pages%></font>ҳ 
��<font color="red"><%=currentpage%></font>ҳ <%if currentpage>1 then%> <font size="2"><a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=currentpage-1%>"><font size="2">[��һҳ]</font></a> 
<%end if%> <%if currentpage<pages then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=currentpage+1%>&search_txt=<%=search_txt%>&px=<%=px%>">[��һҳ]</a> 
<%end if%> <%if currentpage>1 then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=1&search_txt=<%=search_txt%>&px=<%=px%>">[����ҳ]</a> 
<%end if%> <%if currentpage<pages then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=pages%>&search_txt=<%=search_txt%>&px=<%=px%>">[��ĩҳ]</a> 
<%end if%> </font> 
</body>
</html>