<%@ LANGUAGE=VBScript%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "������������ٽ����Ʊ��"
window.close()
</script>
<%Response.End
end if

if sjjh_name="" then
%>
<script language="vbscript">
  alert "���ܽ��룬�㻹û�е�¼����"
  location.href = "../index.asp"
</script>
<%
else
name=sjjh_name
%>
<!--#include file="jhb.asp"-->
<%
	sql="select * from �ͻ� where �ʺ�='" & name & "'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
	%>
<script language="vbscript">
  alert "�㻹û�п����أ�"
  location.href = "jingji.asp"
</script>
<%	   
    else
    jjr=rs("������")
%>
<html>
<head>
<title>�����˰칫��</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>���Ĺ�Ʊ������<%=jjr%>�ڴ�Ϊ������</p><br>
<a href=cun.asp>��Ǯ����Ʊ�ʻ�</a><br><a href=qu.asp>�ӹ�Ʊ�ʻ���Ǯ</a><br><a href=cha.asp>�鿴�����Ʊ�������</a><br><a href=huan.asp>��������</a>  
<br>  
<p align=center><a href=jingji.asp>�뿪</a></p>
</table></table>  

		
</body>
</html>
<%
end if
end if
end if
%>