<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT �Ա� FROM �û� WHERE ����='" & aqjh_name & "'",conn
sex=rs("�Ա�")
rs.close
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>����Ժ</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../../bg.gif">
<br>
<table border=1 bgcolor="#FFFFFF" align=center width=541 cellpadding="0" cellspacing="1" bordercolor="#FFCC00">
<tr bgcolor="#669900"> 
<td height="16" colspan="6" width="533"> 
<p align=center class="p9"><font color="#FF3333"><b><font color="#FFFFFF">��������Ժ</font></b></font></p>
<tr bgcolor=#74E76D bordercolor="#000000"> 
<td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="45"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">��Ů</font></div>
</td>
<td height="15" bgcolor="#FFFFCC" bordercolor="#FF6600" width="49"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">��ò</font></div>
</td>
<td height="15" width="93" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">��������ļ۸�</font></div>
</td>
<td height="15" width="122" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">�����������ӵ�����</font></div>
</td>
<%if sex="��" then%>
<td height="15" width="102" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">����뷿</font></div>
</td>
<%end if%>
<td height="15" width="102" bgcolor="#FFFFCC" bordercolor="#FF6600"> 
<div align="center">
  <p align="center"><font size="2" color="#FF0000">Ϊ������</font></div>
</td>
</tr>
<!--#include file="jiu.asp"-->
<%
sql="SELECT id ,����,��ò�� FROM ��Ů"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr bgcolor=#DEAD63> 
<td bgcolor="#FFFFFF" width="45"> 
<center>
<font size="2"><%=rs("����")%><span class="p9"><font color="#6666FF"></font></span> 
</font> 
</center>
<font size="2"></font> 
<td bgcolor="#FFFFFF" width="49"> 
<div align="center"><font size="2"><%=rs("��ò��")%></font> </div>
<td bgcolor="#FFFFFF" width="93"> 
<div align="center"><font size="2"><%=rs("��ò��")*5%> </font></div>
</td>
<td bgcolor="#FFFFFF" width="122"> 
<div align="center"><font size="2"><%=int(rs("��ò��")/2)%> </font></div>
</td>
<%if sex="��" then%>
<td bgcolor="#FFFFFF" width="102"> 
<center>
<font size="2"><a href=girl.asp?id=<%=rs("id")%>><span class="calen-curr">����뷿</span></a></font> 
</center>
</td>
<%end if%>
<td>
<div align="center"><font size="2"><a href=shushen.asp?id=<%=rs("id")%>>�������</a></font></div>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
connt.close
set connt=nothing
%>
</table>
<div align="center"><br>
  <a href="index.asp">����</a></div><br><br>
<div align="center"><font color="#000000"><b>���ֽ���</b></font></div>
</BODY></HTML> 