<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
fl=Trim(Request.QueryString("fl"))
if fl<>"����" and fl<>"����" and fl<>"����" and fl<>"ͷ��" and fl<>"˫��" and fl<>"װ��" and fl<>"ҩƷ" and fl<>"��ҩ" and fl<>"��Ƭ" then
	Response.Write "<script Language=Javascript>alert('�����ֻ��Ϊ�����֡����֡����ס�ͷ����˫�š�װ�Ρ�ҩƷ����ҩ����Ƭ��������ѡ��');window.close();</script>"
	Response.End
end if
if sjjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('������ʲô����');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM w WHERE c='" & fl & "' and i>0  order by b",conn
%>
<html>
<head>
<title>��Ʒ�����һ�������wWw.happyjh.com��</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center"><font color=yellow><b><%=fl%>����Ʒ</b></font>һ��(װ������Ʒ����)<font face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font></div>
<br>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">ӵ����</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��Ʒ��</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��Ǯ</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">װ��</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��ʽ</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("b")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("a")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("c")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("i")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("e")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("f")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("g")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("h")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("l")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("j")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">ɾ��</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM s WHERE e='" & fl & "' order by b,c",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">ӵ����</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��Ʒ��</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">���� </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">��Ǯ</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">��ʽ</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">����</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("b")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("a")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("e")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("j")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("f")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("g")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("h")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("i")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("n")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("c")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">ɾ��</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<tr align="center"> 
  <td nowrap width="45" height="16"> 
    <div align="center"></div>
  </td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td height="25"> 
      <div align="center">������Բ鿴���Է�����Ʒ��ɾ���Ϳ��Խ����û�����Ʒɾ����</div>
    </td>
  </tr>
</table>
</body>
</html>
