<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../exit.asp'}</script>
<!--#include file="../showchat.asp"-->
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
<!--#include file="dadata.asp"-->
<%
sql="select * from �û� where ����='"&aqjh_name&"'"
set rs=conn.execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ������ˣ����ܽ��μ����"
conn.close
response.end
else
mypai=rs("����")
dj=rs("�ȼ�")
if dj<2 then
conn.close
Response.Write "<script language=javascript>alert('�㻹�ǽ���С�������ܲμ����');window.close();</script>"	
response.end
else
id=request("id")
sql="select * from ����б� where ID=" & id & " and ����='����' and �μ���='����'"
Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then
connt.close
Response.Write "<script language=javascript>alert('��ϯ����Ϊ�㶨�ģ���Ҫ����������������:P');window.close();</script>"	
response.end
else
yhming=rs("�����")
yyou=rs("ӵ����")
nl=rs("����")
tl=rs("����")
jb=rs("���")
lx=rs("����")
ls=rs("����")
if ls<1 then
sql1="delete * from ����б�  where ID=" & id
connt.execute sql1
sql1="delete * from ����� where �����=" & id
connt.execute sql1
connt.close
Response.Write "<script language=javascript>alert('����������Ѿ�����!');history.back();</script>"	
response.end
else
sql="select * from ����� where �μ���='" & aqjh_name & "'  and �����=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
sql2="insert into �����(�μ���,�����) values ('" & aqjh_name & "' , " & id & " )"
connt.execute sql2
sql1="update �û� set ����=����+"&nl&",����=����+"&tl&" where ����='"&aqjh_name&"' "
conn.execute sql1
sql1="update ����б� set ����=����-1 where ID=" & id
connt.execute sql1
conn.close
connt.close
set rs=nothing
says="<font color=red>��С����Ϣ��</font><font color=green>"&aqjh_name&"�μ���"&yyou&"�ڽ�����Ƶ�"&yhming&"�����еġ�"&lx&"����ᣬ�����˲��ٵ�������������</font>"			'��������
call showchat(says)
else
connt.close
set connt=nothing
Response.Write "<script language=javascript>alert('���Ѳμӹ��������ˣ���ô������!');history.back();</script>"
response.end
end if%>
<%end if%>
<%end if%>
<%end if%>
<%end if%>
<html>
<title>�μ����ɹ�...��ӭ���ٿ��ֽ��� http://www.happyjh.com</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
input {border: 1px solid; font-size: 9pt; font-family: "����"; border-color: #000000 solid}
-->
</style>
<body bgcolor="#CCCCCC">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">�μӾ���ɹ�</font></td>
</tr>
<tr bgcolor="#66FF66">
<td align="center" height="157">
<p><font> <img src="jd/ka1.gif" width="204" height="251"></font></p>
</td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20" class="unnamed1"><b><font color="#FF3333">�㾭����һ�����̻��ʣ��ȵ�����һ��ģ������������ <%=nl%>������ <%=tl%>,�쵽��𣿣��������</font></b></td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20"><b><font color="#FF3333"><input  onClick="javascript:window.document.location.href='jd.asp'" value="�� ��" type=button name="back">  <INPUT onclick=window.close() type=button value=�ر�>  </font></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>