<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>��Ա�б�</title>
<LINK href=css/css.css type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000"><a href="javascript:this.location.reload()">ˢ��</a> 
  <br>
��л��Щ���Ѷ����ǽ����Ĵ���֧�֣������Ա��<a href="../chat/hy.asp" target="_blank">����</a><br>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='����' and ��Ա�ȼ�>0 order by ��Ա���� ",conn
%>
<table width="485" border="1" style="border-collapse: collapse" bordercolor="#B8AF86" cellspacing="0" cellpadding="0" bgcolor="f2f2ea" cellspacing=0 cellpadding=0 height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center">ID</div>
          </td>
          <td width="47" height="10"> 
            <div align="center">����</div>
          </td>
          <td width="25" height="10"> 
            <div align="center">�Ա�</div>
          </td>
          <td width="63" height="10"> 
            <div align="center">����</div>
          </td>
          <td width="54" height="10"> 
            <div align="center">���</div>
          </td>
          <td width="82" height="10"> 
            <div align="center">��Ա��</div>
          </td>
          <td width="88" height="10"> 
            <div align="center">��Ա��������</div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center">�����ȼ�</div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center">��¼</div>
          </td>
        </tr>
<%
dengji1=0
dengji2=0
dengji3=0
dengji4=0
dengji5=0
do while not rs.bof and not rs.eof
Select Case rs("��Ա�ȼ�")
Case 1
	dengji1=dengji1+1
case 2
	dengji2=dengji2+1
case 3
	dengji3=dengji3+1
case 4
	dengji4=dengji4+1
case 5
	dengji5=dengji5+1
case 6
	dengji6=dengji6+1
case 7
	dengji7=dengji7+1
case 8
	dengji8=dengji8+1
end select
%>
        <tr height="20"> 
          <td width="28"> 
            <div align="center"><%=rs("ID")%></div>
          </td>
          <td width="47"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("����")%>"><font color="#FF9900"><%=rs("����")%></a></div>
          </td>
          <td width="25"> 
            <div align="center"><%=rs("�Ա�")%></div>
          </td>
          <td width="63"> 
            <div align="center"><%=rs("����")%></div>
          </td>
          <td width="54""> 
            <div align="center"><%=rs("���")%></div>
          </td>
          <td width="82"> 
            <div align="center"><%=rs("��Ա�ȼ�")%></div>
          </td>
          <td width="88"> 
            <div align="center"><%=rs("��Ա����")%></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><%=rs("�ȼ�")%></div>
          </td>
          <%end if%>
          <td> 
            <div align="center"><%=rs("times")%></div>
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
<div align="center">һ����Ա��<%=dengji1%> ������Ա:<%=dengji2%> ������Ա��<%=dengji3%> �ļ���Ա��<%=dengji4%> �弶��Ա��<%=dengji5%> ������Ա:<%=dengji6%> �߼���Ա��<%=dengji7%> �˼���Ա��<%=dengji8%><br>
��Ա������<%=(dengji1+dengji2+dengji3+dengji4+dengji5+dengji6+dengji7+dengji8)%>��</div></body></html>
