<% if Session("myName")="" then
  response.write("SORRY����û��<a href=../index.htm>��¼</a>ϵͳ��")
else%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM ���� where  ����='һ�㷿��'"
Set Rs=conn.Execute(sql)
name=session("myname")
%>
<html>
<LINK 
href="../pic/css.css" rel=stylesheet>
<body bgcolor="#FFFFFF">
<table border=1 bgcolor="#FFFFFF" align=center width=54% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099">��ӭ</font><%=name%><font color="#990099">���ٷ������ģ�˭�������Լ��ļң���������������ĸ���Ҫ��(������������ֻ�ܵõ�ԭ����1/5�ļ۸�)</font></marquee></font></p>
  <tr bgcolor="#006633"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="1" cellspacing="1" cellpadding="1" bordercolor="#6633FF">
        <tr> 
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu.asp><font color="#FF3333">һ�㷿��</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu3.asp><font color="#FF3333">�߼���Ԣ</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu4.asp><font color="#FF3333">��԰��</font></a></div>
          </td>
          <td bgcolor="#9999FF"> 
            <div align="center"><a href=fangwu5.asp><font color="#FF3333">��������</font></a></div>
          </td>
        </tr>
      </table>
  <tr bgcolor=#74E76D bordercolor="#000000"> 
    <td height="15" width="14%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">��������</font></div>
    </td>
    <td height="15" width="9%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">�ۼ�</font></div>
    </td>
    <td height="15" width="11%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">����</font></div>
    </td>
    <td height="15" colspan="2" bgcolor="#FF9933" bordercolor="#FF6600">
      <div align="center"><font size="2" color="#FF3333">���ƺ���</font></div>
    </td>
    <td height="15" width="19%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF3333">����/����</font></div>
    </td>
    <td height="15" width="25%" bgcolor="#FF9933" bordercolor="#FF6600"> 
      <div align="center"><font size="2" color="#FF0000">����</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor=#DEAD63> 
    <td width="14%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("����")%> 
      </center>
      </font> 
    <td width="9%" bgcolor="#FFFFFF" class="calen-curr" height="15"><%=rs("�ۼ�")%> 
      <div align="center"></div>
    <td width="11%" bgcolor="#FFFFFF" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("λ��")%></font></div>
    </td>
    <td colspan="2" bgcolor="#FFFFFF" class="ct-def4" height="15"><%=rs("�ֵ�")%> 
      <font color=red><%=rs("id")%></font><font size="2">��</font></td>
    <td width="19%" bgcolor="#FFFFFF" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("����")%> / </font><font color="ff00ff"><%=rs("����")%></font></div>
    </td>
    <td width="25%" bgcolor="#FFFFFF" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("����")<>name and rs("����")="��" then%><a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">����</font></a><%end if%><%if rs("����")=sjjh_name then%><a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>����</font></a><%end if%><%if rs("����")<>"��" and rs("����")<>name then%>�÷�������<%end if%> 
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 
conn.close 
%> 
</table>
</html>
<%end if%>