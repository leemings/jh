<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
tiaojian=Request.Form("tiaojian")
show=trim(Request.Form("show"))
fs=int(Request.form("seekfs"))
%>
<html>
<head>
<title>�û����ϲ鿴����</title>
<link rel="stylesheet" href="../STYLE.CSS">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body   leftmargin="0" background='../chatroom/bg1.gif' >
<p align="center"> <a href="javascript:this.location.reload()"> <font color="#008000" face="����" size="4">���в�������ˢ��</font> 
</a> 
  <br>
  ��һЩ�������������ˣ������������޸ģ�<br>
  <font color="#FF0000"><b><%=tiaojian%></b></font> <br>
 <%
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
select case fs
case 0
	if show<>"" then
		rs.open "SELECT * FROM �û� where "& tiaojian &" order by -"&show,conn
	else
		rs.open "SELECT * FROM �û� where "& tiaojian &" order by �ȼ�",conn
	end if
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
 align="center">
  <tr align="center">
    <td  width="500" height="26"></td>
</tr>
<tr align="center">
<td>
      <table width="485" border="1" cellpadding="0" cellspacing="1" height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center"><font color="#000000">ID</font></div>
          </td>
          <td width="47" height="10"> 
           <div align="center"><font color="#000000">�ʺ�</font></div>
          </td>
          <td width="47" height="10"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <td width="25" height="10"> 
            <div align="center"><font color="#000000">�Ա�</font></div>
          </td>
          <td width="63" height="10"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <td width="54" height="10"> 
            <div align="center"><font color="#000000">���</font></div>
          </td>
          <td width="82" height="10"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <td width="58" height="10"> 
            <div align="center"><font color="#000000">�ȼ�</font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center"><font color="#000000"><b><%=show%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center"><font color="#000000">��Ա</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="28" height="30"> 
            <div align="center"><font color="#000000"><%=rs("ID")%></font></div>
          </td>
             <td width="28" height="30"> 
            <div align="center"><font color="#000000"><%=rs("�ʺ�")%></font></div>
          </td>
          <td width="47" height="30"> 
            <div align="center"><a href="SHOWUSER.asp?search=<%=rs("����")%>"><font color="#0000FF"><%=rs("����")%></font></a></div>
          </td>
          <td width="25" height="30"> 
            <div align="center"><font color="#000000"><%=rs("�Ա�")%></font></div>
          </td>
          <td width="63" height="30"> 
            <div align="center"><font color="#000000"><%=rs("����")%></font></div>
          </td>
          <td width="54" height="30"> 
            <div align="center"><font color="#000000"><%=rs("���")%></font></div>
          </td>
          <td width="82" height="30"> 
            <div align="center"><font color="#000000"><%=rs("����")%></font></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#000000"><%=rs("�ȼ�")%></font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><font color="#000000"><b><%=rs(show)%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="30"> 
            <div align="center"><font color="#000000"><%=rs("��Ա")%></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
</td>
</tr>
<tr align="center">
    <td  width="500" height="2"><font color="#000000">��Ϸ�ɰ�Ȩ����</font></td>
</tr>
</table>
<div align="center"><font color="#000000">��������:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">��</font> 
</div>
<% case 1
sql="SELECT * FROM ��Ʒ where "& tiaojian &" order by ����,����"
Set Rs=conn.Execute(sql)
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
 align="center">
  <tr align="center"> 
    <td  width="500" height="26"></td>
  </tr>
  <tr align="center"> 
    <td> 
      <table width="485" border="1" cellpadding="0" cellspacing="1" height="13">
        <tr> 
          <td width="26" height="7"> 
            <div align="center"><font color="#000000">ID</font></div>
          </td>
          <td width="41" height="7"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <td width="52" height="7"> 
            <div align="center"><font color="#000000">������</font></div>
          </td>
          <td width="32" height="7"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <td width="110" height="7"> 
            <div align="center"><font color="#000000">��Ч</font></div>
          </td>
          <td width="34" height="7"> 
            <div align="center"><font color="#000000">װ��</font></div>
          </td>
          <td width="34" height="7"> 
            <div align="center"><font color="#000000">�۸�</font></div>
          </td>
          <td width="58" height="7"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
          <%if show<>"" then%>
                    <td width="58" height="7"> 
            <div align="center"><font color="#000000"><%=show%></font></div>
          </td>
          <%end if%>
          <td width="58" height="7"> 
            <div align="center"><font color="#000000">����</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="26" height="30"> 
            <div align="center"><font color="#000000"><%=rs("ID")%></font></div>
          </td>
          <td width="41" height="30"> 
            <div align="center"><font color="#000000"><%=rs("����")%></font></div>
          </td>
          <td width="52" height="30"> 
            <div align="center"><font color="#000000"><%=rs("������")%></font></div>
          </td>
          <td width="32" height="30"> 
            <div align="center"><font color="#000000"><%=rs("����")%></font></div>
          </td>
          <td width="110" height="30"> 
            <div align="center"><font color="#000000"><%=rs("��Ч")%></font></div>
          </td>
          <td width="34" height="30"> 
            <div align="center"><font color="#000000"><%=rs("װ��")%></font></div>
          </td>
          
          <td width="53" height="30"> 
            <div align="center"><b><font color="#000000"><%=rs("�۸�")%></font></b></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#000000"><%=rs("����")%></font></div>
          </td>
          <%if show<>"" then%>
            <td width="58" height="30"> 
            <div align="center"><font color="#000000"><%=rs(show)%></font></div>
          </td>
          <%end if%>
          <td width="59" height="30"> 
            <div align="center"><font size="-1"><a href="Chgcur2.asp?id=<%=rs("id")%>&uname=<%=rs("������")%>&ucur=<%=rs("����")%>"><font color="#0000FF">ɾ��</font></a></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
'conn.close
 case 2
sql="SELECT * FROM ���� where "& tiaojian &" order by �ȼ�"
Set Rs=conn.Execute(sql)
%>
      </table>
    </td>
  </tr>
  <tr align="center"> 
    <td width="500" height="2"><font color="#000000">ħ�ý�����Ȩ����</font></td>
  </tr>
</table>
<div align="center"><font color="#000000">��������:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000"></font> 
  <br>
  <table border="0" width="500" cellspacing="0" cellpadding="0"
 align="center">
    <tr align="center"> 
      <td width="500" height="26"></td>
    </tr>
    <tr align="center"> 
      <td> 
        <table width="485" border="1" cellpadding="0" cellspacing="1" height="13">
          <tr> 
            <td width="32" height="12"> 
              <div align="center"><font color="#000000">ID</font></div>
            </td>
            <td width="64" height="12"> 
              <div align="center"><font color="#000000">��ʽ</font></div>
            </td>
            <td width="71" height="12"> 
              <div align="center"><font color="#000000">����</font></div>
            </td>
            <td width="61" height="12"> 
              <div align="center"><font color="#000000">��ʽ�ȼ�</font></div>
            </td>
            <td width="233" height="12"> 
              <div align="center"><font color="#000000">˵��</font></div>
            </td>
                        <td width="61" height="12"> 
              <div align="center"><font color="#000000">�ľ���</font></div>
            </td>
                        <td width="61" height="12"> 
              <div align="center"><font color="#000000">������</font></div>
            </td>
            
            <td width="55" height="12"> 
              <div align="center"><font color="#000000">��������</font></div>
            </td>
            <%if show<>"" then%>
                        <td width="55" height="12"> 
              <div align="center"><font color="#000000"><%=show%></font></div>
            </td>
            <%end if%>
            <td width="53" height="12"> 
              <div align="center"><font color="#000000">����</font></div>
            </td>
          </tr>
          <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
          <tr> 
            <td width="32" height="16"> 
              <div align="center"><font color="#000000"><%=rs("ID")%></font></div>
            </td>
            <td width="64" height="16"> 
              <div align="center"><font color="#000000"><%=rs("��ʽ")%></font></div>
            </td>
            <td width="71" height="16"> 
              <div align="center"><font color="#000000"><%=rs("����")%></font></div>
            </td>
            <td width="61" height="16"> 
              <div align="center"><font color="#000000"><%=rs("�ȼ�")%></font></div>
            </td>
            <td width="233" height="16"> 
              <div align="center"><font color="#000000"><%=rs("����˵��")%></font></div>
            </td>
                        <td width="61" height="16"> 
              <div align="center"><font color="#000000"><%=rs("���ľ���")%></font></div>
            </td>
                        <td width="61" height="16"> 
              <div align="center"><font color="#000000"><%=rs("��������")%></font></div>
            </td>
            <td width="55" height="16"> 
              <div align="center"><font color="#000000"><%=rs("��������")%></font></div>
            </td>
            <%if show<>"" then%>
              <td width="55" height="12"> 
            <div align="center"><font color="#000000"><%=rs(show)%></font></div>
          </td>
            <%end if%>
            <td width="53" height="16"> 
              <div align="center"><font size="-1"><a href="chgwu2.ASP?id=<%=rs("id")%>&uname=<%=rs("����")%>"><font color="#0000FF">ɾ��</font></a></font></div>
            </td>
          </tr>
          <%
rs.movenext
loop
conn.Close
set conn=nothing
%>
        </table>
      </td>
    </tr>
    <tr align="center"> 
      <td  width="500" height="5"><font color="#000000">ħ�ý�������Ϸ�ɰ�Ȩ����</font></td>
    </tr>
  </table>
  <div align="center"><font color="#000000">��������:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000"></font> 
  </div>
  <br>
</div>
<%end select%>
</body>
</html>
