<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Application("aqjh_usermdb")
Set Rs = Server.CreateObject("ADODB.Recordset")
Sql_1="SELECT * from b WHERE b = '���' or b = '���' ORDER BY b ASC"
'Sql_1="SELECT * from b ORDER BY b ASC"
Rs.Open Sql_1,Conn,0,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function wp_gm(wp_name) {  //reloads the window if Nav4 resized
var num
num=parseInt(window.prompt("��ӭ������Ʒ���:"+wp_name+" ���ȷ������",0));
if (isNaN(num) || num<0 || num>100){
window.status="�������������99��"
form_wp.post.value=''
alert("�������������99��")
}else{
form_wp.wp_name.value=wp_name
form_wp.wp_num.value=num
//form_wp.post.value='post'
document.form_wp.submit()
}
}
function checkdata(){
if( form_wp.post.value != 'post' ) {
return false;
}
return true;
}
//-->
</script>
<link href="base.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {color: #0000FF}
-->
</style>
</head>
<body> 
<div id="div_gm" style="position:absolute; left:0px; top:0px; width:154px; height:80px; z-index:1; visibility: hidden;"> 
  <table border="0" cellspacing="0" cellpadding="4"> 
    <form action="wpcz.asp"  method="post" name="form_wp" target="rightFrame" id="form_wp" onsubmit="return checkdata()"> 
      <tr> 
        <td><input type="text" name="wp_name" /></td> 
      </tr> 
      <tr> 
        <td><input type="text" name="wp_num" /></td> 
      </tr> 
      <tr> 
        <td><input type="submit" name="Submit" value="�ύ" /> 
          <input name="post" type="hidden" value="" /> </td> 
      </tr> 
    </form> 
  </table> 
</div> 
�ֽ������۵Ķ�..
<table width="100%" border="0" align="right" cellpadding="0" cellspacing="3"> 
  <tr align="center" bgcolor="#FF9900"> 
    <td>����</td> 
    <td>���</td> 
    <td>˵��</td> 
    <td>ͼƬ</td> 
    <td>�۸�</td> 
    <td>��</td> 
  </tr> 
  <% While (NOT Rs.EOF) 
  wp_image=(Rs.Fields.Item("i").Value)
  if isnull(wp_image) or wp_image="" or wp_image="��" then  wp_image="0.gif"
  wp_j=(Rs.Fields.Item("h").Value)
  %> 
  <tr align="center" bgcolor="#FFCC99"> 
    <td><%=(Rs.Fields.Item("a").Value)%></td> 
    <td><%=(Rs.Fields.Item("b").Value)%></td> 
    <td><%=(Rs.Fields.Item("c").Value)%></td> 
    <td><img style = "FILTER: chroma(color:#FFFFFF)" src="images/<%= wp_image %>" width="30" height="30" /></td> 
    <td><%= wp_j %></td> 
    <td><a class="style1" onclick="wp_gm('<%=(Rs.Fields.Item("a").Value)%>');">����</a>  
  </tr> 
  <% Rs.MoveNext()
Wend
%> 
</table> 
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
</body>
</html>

