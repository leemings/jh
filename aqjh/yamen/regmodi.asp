<%@ LANGUAGE=VBScript codepage ="936" %>
<%

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("aqjh_usermdb")
action=aqjh_name
sql="Select  * from �û� where ����='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>

<head>
<title>���������޸�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: ���� }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: ����; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body bgcolor="#BABABA" topmargin="0" leftmargin="0" background=../bg.gif>
<form method="POST" action="regmodi1.asp">
<table border="1" width="429" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
<tr bgcolor="#000000">
<td colspan="8" height="33">
<div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("����")%></font></b><font
size="2" color="#FFFFFF">�����Ľ�������</font></font></font> </div>
</td>
</tr>
<tr bgcolor="#cccccc">
<td width="36" height="12"><font size="2" class="c">�� ��:</font></td>
<td colspan="2" height="12"><%=rs("����")%></td>
<td height="12" colspan="2">����:
<input type="text"
name="diqu" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("����")%>">
</td>
<td height="12" width="66">
<div align="center">�� ��:</div>
</td>
<td height="12" colspan="2"> <font color="#0000FF" size="2">
<%if rs("�ȼ�")<5  then response.write("����է��")
if rs("�ȼ�")>=5 and rs("�ȼ�")<10  then response.write("��������")
if rs("�ȼ�")>=10 and rs("�ȼ�")<15  then response.write("С�гɾ�")
if rs("�ȼ�")>=15 and rs("�ȼ�")<20  then response.write("�����Ժ�")
if rs("�ȼ�")>=20 and rs("�ȼ�")<35  then response.write("�д�����")
if rs("�ȼ�")>=35 and rs("�ȼ�")<45  then response.write("һ������")
if rs("�ȼ�")>=45 and rs("�ȼ�")<55  then response.write("��������")
if rs("�ȼ�")>=55 and rs("�ȼ�")<65  then response.write("��������")
if rs("�ȼ�")>=65 and rs("�ȼ�")<75  then response.write("������ʤ")
if rs("�ȼ�")>=100 then response.write("����")
%>
</font></td>
</tr>
<tr bgcolor="#cccccc">
<td width="36" height="20"><font size="2" class="c">�� ��:</font></td>
<td height="20" width="38"><font color="#000000" size="2">
<%if rs("�Ա�")="��"  then response.write("˧��")
if rs("�Ա�")="Ů"  then response.write("��Ů")
%>
</font></td>
<td height="20" width="35">����:</td>
<td height="20" nowrap width="71">
<input type="text"
name="nianling" size="2"
style="font-family: Tahoma; font-size: 12px"
maxlength="3" value="<%=rs("����")%>">
�� </td>
<td height="20" width="40">״̬:</td>
<td height="20" width="66"><%=rs("״̬")%></td>
<td height="20" width="36">ʦ��:</td>
<td height="20" width="73"><%=rs("ʦ��")%></td>
</tr>
<tr bgcolor="#cccccc">
<td width="36"><font size="2" class="c">Email:</font></td>
<td colspan="3" nowrap>
<input type="text"
name="email" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("����")%>">
</td>
<td width="40">OIcq:</td>
<td width="66">
<input type="text"
name="oicq" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9" value="<%=rs("oicq")%>">
</td>
<td width="36">ְҵ:</td>
<td width="73"><%=rs("ְҵ")%></td>
</tr>
</table>
<div align="left"></div>
  <table border="1"
width="424" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
    <tr bgcolor="#cccccc"> 
      <td colspan="5" height="20"> 
        <div align="center"> <font size="2">�� �� �� ��</font> </div>
      </td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="2">�� ��</td>
      <td width="169" height="2"><%=rs("����")%> ��</td>
      <td width="56" height="2">�����ˣ�</td>
      <td height="2" width="50"><%=rs("������")%></td>
      <td height="2" width="84">��Ա�ѣ�<%=rs("��Ա��")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� �</td>
      <td width="169" height="24"><%=rs("���")%> ��</td>
      <td width="56" height="24">�� �� ��</td>
      <td height="24" colspan="2"><%=rs("����")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� ����</td>
      <td width="169" height="24"><%=rs("�书")%></td>
      <td width="56" height="24">�� �� ��</td>
      <td height="24" colspan="2"><%=rs("����")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� ����</td>
      <td width="169"><%=rs("����")%></td>
      <td width="56">�� �� ��</td>
      <td colspan="2"><%=rs("����")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� ����</td>
      <td width="169"><%=rs("����")%></td>
      <td width="56">�� �� ��</td>
      <td colspan="2"><%=rs("����")%> </td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� ����</td>
      <td width="169"><%=rs("����")%></td>
      <td width="56">�� �� ��</td>
      <td colspan="2"><%=rs("����")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">�� ż��</td>
      <td width="169"><%=rs("��ż")%></td>
      <td width="56">Ӯ / �ģ�</td>
      <td colspan="2"><%=rs("Ӯ����")%> / <%=rs("�Ĵ���")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="2">�� ����</td>
      <td width="169" height="2"> 
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
      <td width="56" height="2">�� ����</td>
      <td width="50" height="2"><%=rs("�ȼ�")%>��</td>
      <td width="84" height="2">������:<%=rs("grade")%>��</td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td colspan="5" height="199"> 
        <table width="421" border="1" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="196" width="237"> 
              <div align="center"><font color="#000000" size="2"> </font> <font color="#000000"><img src="<%=rs("ͷ��")%>"></font><br>
                <br>
                ͷ��url��ַ�� <br>
                <input type="text"
name="touxiang" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="100" value="<%=rs("ͷ��")%>">
                <br>
                ����ͷ���url��ַ�������޸��Լ�<br>
                ��ͷ�������������ס��ݷ����е�<br>
                ͷ�������Լ��ģ�ͷ��Ҫ̫�󣺽���<br>
                ���ã�200*130��С </div>
            </td>
            <td height="196" width="178" align="left" valign="top"> 
              <div align="center">����ǩ����д����������˵�Ļ���<span class="txt"> 
                <textarea name="qianming" cols="30" style="font-family: Tahoma; font-size: 12px" rows="20"><%=rs("ǩ��")%></textarea>
                </span></div>
            </td>
          </tr>
        </table>
        <div align="right"></div>
        <div align="right"></div>
      </td>
    </tr>
  </table>
<div align="center"><br>
<input type="submit" value="���������޸�" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</div>
</form>
<div align="center">
<br>
</div>
</body>

</html>
<%
rs.close
set rs=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%>