<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>�����б�</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font color="#0000FF">�����б�<br>
  <br>
  ��ʾ��0���ʾ����1���ʾ��ֹ��ֻ����վ�������޸�������������,<br>
  �޸���������������������Ŀ¼�ļ���global.asa�����ٸĻ����Ż���Ч�� <br><font color=red>ע�⣺���������npc�ķ��ŷ��䡢��Թ����Ƕ��������жϵķ��䣬����������벻Ҫ���κθĶ���</p>
<table width="90%" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" align="center" bgcolor="#006699">
  <tr> 
    <td colspan="8" height="11" bgcolor="#336633"> 
      <p align="center"><font color="#FFFFFF">�� �� �� �� �� ��</p>
    </td>
    <td height="11" bgcolor="#336633"> 
      <div align="center"><font color="#000000"> <a href="addroom.asp"><font color="#FFFFFF">��������</a></div>
    </td>
  </tr>
  <tr> 
    <td width="104" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">������</div>
    </td>
    <td width="64" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">����</div>
    </td>
    <td width="68" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">����</div>
    </td>
    <td width="74" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">�¼�</div>
    </td>
    <td width="56" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">����</div>
    </td>
    <td width="60" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">��Ƭ</div>
    </td>
    <td width="72" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">�Ĳ�</div>
    </td>
    <td width="84" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">����</div>
    </td>
    <td width="89" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">���� </div>
    </td>
  </tr>
  <%
sql="SELECT * FROM r"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr> 
    <td width="104" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("a")%> </div>
    </td>
    <td width="64" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("b")%></div>
    </td>
    <td width="68" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("f")%></div>
    </td>
    <td width="74" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("g")%></div>
    </td>
    <td width="56" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("h")%></div>
    </td>
    <td width="60" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("i")%></div>
    </td>
    <td width="72" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("j")%></div>
    </td>
    <td width="84" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("c")%></div>
    </td>
    <td width="89" height="23"> 
      <div align="center"> <a href="modiroom.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF">�޸�</a> 
        <%if rs("id")<>1 then %>
        <a href="delroom.asp?id=<%=rs("id")%>"><font color="#FFFFFF">ɾ��</a> 
        <%end if%>
      </div>
    </td>
  </tr>
  <tr> 
    <td colspan="9" height="23"><font color="#FFFF00">���ջ���:<font color="#FFFF00" ><%=rs("k")%></td>
  </tr>
  <%if rs("c")<>0 then %>
  <tr> 
    <td colspan="9" height="16"> 
      <table width="100%" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#006699">
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font color="#000000">����˵��</div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#CCCCFF"> <font color="#000000">���Ʊ��ʽ</div>
          </td>
        </tr>
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font color="#0000FF" ><%=rs("d")%></div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#0000FF" ><%=rs("e")%></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <%
 end if
rs.movenext
loop
conn.close
%>
</table>
<p align=center><a href=room_reset.asp?rst=1>���������б�</p>
</body></html>