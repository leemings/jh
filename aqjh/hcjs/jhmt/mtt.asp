<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
username=Request.Form("search")
if username=Application("aqjh_user") then
   Response.Write "<script language=javascript>alert('���ܵ���վ�������ϣ�');history.back();</script>"
end if
rs.open "Select  * from �û� where ����='"& username &"'",conn
if rs.EOF or rs.BOF then
 Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
   rs.close
   set rs=nothing	
   conn.close
   set conn=nothing
   Set conn=Server.CreateObject("ADODB.CONNECTION")
   Set rs=Server.CreateObject("ADODB.RecordSet")
   conn.open Application("aqjh_usermdb")
   rs.open "Select  * from �û� where ����='"& aqjh_name &"'",conn,2,2
   if rs("����")<1000000 or rs("���")<10 then
     Response.Write "<script language=javascript>alert('��������100����Ҳ���10ö');history.back();</script>"
   else
      conn.execute "update �û� set ����=����-1000000,���=���-10 where ����='" & aqjh_name &"'"
      rs.close
      set rs=nothing	
      conn.close
      set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
      rs.open "Select  * from �û� where ����='"& username &"'",conn
   end if
%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language=javascript>window.moveTo(100,50);</script>
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: ���� }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: ����; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>
<body style="BACKGROUND-COLOR: buttonface; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; BORDER-RIGHT: medium none; BORDER-TOP: medium none; PADDING-BOTTOM: 6px; PADDING-LEFT: 6px; PADDING-RIGHT: 6px; PADDING-TOP: 6px" marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" LANGUAGE="javascript"  >
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center">
  <tr bgcolor="#000000">
<td colspan="8" height="33">
<div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF">ID:</font><font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("id")%>
</font></b></font></font><font color="#FFFFFF"><%=rs("����")%></font></b><font
size="2" color="#FFFFFF">�����Ľ�������</font></font></font> </div>
</td>
</tr>
<tr >
<td width="36" height="12"><font size="2" class="c">�� ��:</font></td>
<td colspan="2" height="12"><%=rs("����")%></td>
<td height="12" colspan="2">����:<%=rs("����")%></td>
<td height="12" width="49">
<div align="center">�� ��:</div>
</td>
<td height="12" colspan="2"> <font color="#0000FF" size="2">
<%
if aqjh_name=username then
zddj=-1
else
zddj=rs("�ȼ�")
end if
if rs("�ȼ�")<5  then response.write("����է��")
if rs("�ȼ�")>=5 and rs("�ȼ�")<10  then response.write("��������")
if rs("�ȼ�")>=10 and rs("�ȼ�")<15  then response.write("С�гɾ�")
if rs("�ȼ�")>=15 and rs("�ȼ�")<20  then response.write("�����Ժ�")
if rs("�ȼ�")>=20 and rs("�ȼ�")<35  then response.write("�д�����")
if rs("�ȼ�")>=35 and rs("�ȼ�")<45  then response.write("һ������")
if rs("�ȼ�")>=45 and rs("�ȼ�")<55  then response.write("��������")
if rs("�ȼ�")>=55 and rs("�ȼ�")<65  then response.write("��������")
if rs("�ȼ�")>=65 and rs("�ȼ�")<75  then response.write("������ʤ")
if rs("�ȼ�")>=75 and rs("�ȼ�")<80  then response.write("�����ɵ�")
if rs("�ȼ�")>=80 and rs("�ȼ�")<85  then response.write("С��")
if rs("�ȼ�")>=85 and rs("�ȼ�")<90  then response.write("����")
if rs("�ȼ�")>=90 and rs("�ȼ�")<95  then response.write("���۴���")
if rs("�ȼ�")>=95 and rs("�ȼ�")<100  then response.write("����")
if rs("�ȼ�")>=100 then response.write("����")
%>
</font></td>
</tr>
<tr >
<td width="36" height="14"><font size="2" class="c">�� ��:</font></td>
<td height="14" width="33"><font color="#000000" size="2">
<%if rs("�Ա�")="��"  then response.write("˧��")
if rs("�Ա�")="Ů"  then response.write("��Ů")
%>
</font></td>
<td height="14" width="37">����:</td>
<td height="14" nowrap width="96"><%=rs("����")%> ��</td>
<td height="14" width="41">״̬:</td>
<td height="14" width="49"><%=rs("״̬")%></td>
<td height="14" width="34">ʦ��:</td>
<td height="14" width="69"><%=rs("ʦ��")%></td>
</tr>
<tr >
<td width="36"><font size="2" class="c">Email:</font></td>
<td colspan="3" nowrap><%=rs("����")%> </td>
<td width="41">OIcq:</td>
<td width="49"><%=rs("oicq")%></td>
<td width="34">ְҵ:</td>
<td width="69"><%=rs("ְҵ")%></td>
</tr>
</table>
<div align="left"></div>
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center">
  <tr > 
    <td colspan="5" height="20"> 
      <div align="center"> <font size="2">�� �� �� ��</font> </div>
    </td>
  </tr>
  <tr > 
    <td width="74" height="25">�� ��</td>
    <td height="25" width="134"> 
      <%if aqjh_jhdj>zddj then%>
      <%=rs("����")%> �� 
      <%else%>
      �ȼ����� 
      <%end if%>
    </td>
    <td width="74" height="25">�����ˣ�</td>
    <td height="25" width="47"> <%=rs("������")%> </td>
    <td height="25" width="93">�𿨣� 
      <%if aqjh_jhdj>zddj then%>
      <font color="#0000FF"><%=rs("��Ա��")%>Ԫ</font> 
      <%else%>
      �ȼ����� 
      <%end if%>
    </td>
  </tr>
  <tr > 
    <td width="74" height="20">�� �</td>
    <td height="24" width="134"> 
      <%if aqjh_jhdj>zddj then%>
      <%=rs("���")%> �� 
      <%else%>
      �ȼ����� 
      <%end if%>
    </td>
    <td width="74" height="24">�� �£�</td>
    <td height="24" colspan="2"> <%=rs("����")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">�� ����</td>
    <td height="24" width="134"> <%=rs("�书")%></td>
    <td width="74" height="24">�� ����</td>
    <td height="24" colspan="2"> <%=rs("����")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">�� ����</td>
    <td width="134"> <%=rs("����")%> </td>
    <td width="74">�� ����</td>
    <td colspan="2"> <%=rs("����")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">�� ����</td>
    <td width="134"> <%=rs("����")%> </td>
    <td width="74">�� �ɣ�</td>
    <td width="47"><%=rs("����")%> </td>
    <td width="93">����:<%=rs("����")%></td>
  </tr>
  <tr > 
    <td width="74" height="20">�� ����</td>
    <td width="134"> <%=rs("����")%> </td>
    <td width="74">�� �ݣ�</td>
    <td colspan="2"> <%=rs("���")%> </td>
  </tr>
<tr > 
    <td width="74" height="20">�� �ң�</td>
    <td width="134"> <%=rs("����")%> </td>
    <td width="74">ְ λ��</td>
    <td colspan="2"> <%=rs("ְλ")%> </td>
  </tr>
<tr > 
    <td width="74" height="20">�� �壺</td>
    <td width="134"> <%=rs("����")%> </td>
    <td width="74">�� ����</td>
    <td colspan="2"> <%=rs("����")%> </td>
  </tr>
  <tr > 
    <td width="74" height="27">�� ż��</td>
    <td width="134" height="27"> <%=rs("��ż")%> <font color="#0000FF">������:<%=rs("������")%></font></td>
    <td width="74" height="27">Ӯ / �ģ�</td>
    <td colspan="2" height="27"><%=rs("Ӯ����")%> / <%=rs("�Ĵ���")%></td>
  </tr>
  <tr > 
    <td width="74" height="20">�ĳ��ȼ���</td>
    <td width="134"> 
      <%if rs("ӮǮ")<100  then response.write("�ĳ�����")
if rs("ӮǮ")>=2000 and rs("ӮǮ")<10000  then response.write("�ĳ�����")
if rs("ӮǮ")>=10000 and rs("ӮǮ")<50000  then response.write("������ͽ")
if rs("ӮǮ")>=50000 and rs("ӮǮ")<100000  then response.write("�ĳ�����")
if rs("ӮǮ")>=100000 and rs("ӮǮ")<300000  then response.write("ְҵ����")
if rs("ӮǮ")>=300000 and rs("ӮǮ")<800000  then response.write("����")
if rs("ӮǮ")>=800000 and rs("ӮǮ")<1500000  then response.write("����")
if rs("ӮǮ")>=1000000 then response.write("����")
%>
    </td>
    <td width="74">�����ȼ���</td>
    <td width="47"><%=rs("�ȼ�")%>��</td>
    <td width="93">����ȼ�:<%=rs("grade")%>��</td>
  </tr>
</table></body></html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
end if
%>