<%if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
dim conn,rs
%>

<html>
<head>
<title>�ύ״ֽ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: ����; font-size: 9pt}body {  font-family: ����; font-size: 9pt}select {  font-family: ����; font-size: 9pt}A {text-decoration: none; font-family: "����"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "����"; font-size: 9pt} .big {  font-family: ����; font-size: 12pt}.txt {  font-family: "����"; font-size: 11pt}
--></style>
<script LANGUAGE="javascript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.topic.value=="")
	{
	  alert("״ֽ��Ŀû���")
	  document.FrmAddLink.topic.focus()
	  return false
	 }
else if(document.FrmAddLink.name.value=="")
	{
	  alert("״�����û���")
	  document.FrmAddLink.name.focus()
	  return false
	 }
else if(document.FrmAddLink.play.value=="")
	{
	  alert("�Ա����Ҫ��û���")
	  document.FrmAddLink.play.focus()
	  return false
	 }
else if(document.FrmAddLink.text.value=="")
	{
	  alert("״��û���")
	  document.FrmAddLink.text.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body oncontextmenu=self.event.returnValue=false text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" bgcolor="#3a4b91">
<table width=590 border=0 cellspacing=0 cellpadding=0 align=center height="19" bgcolor="#0066CC">
  <tr>
    <td height="26" width="86%"><b>�ύ״ֽ</b></td>
    <td align=right height="26" width="14%"> 
      <div align="center"><a onClick='javascript:history.back()'>�� 
        ��</a></div>
    </td>
  </tr>
</table>
<form action=newplan.asp method=post  LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
  <table border=0 cellspacing=0 cellpadding=0 align=center width="590">
    <tr bgcolor="#0066CC"> 
      <td height="26" colspan="2"> <b>״ֽ��Ŀ:</b> 
        <input name=topic size=28 maxlength="12">     
        �12��</td>     
    </tr>     
    <tr>      
      <td height="17" colspan="2"><b>״�����: <input name=name size=28 maxlength="5">   
        ��д��������������������</b></td>    
    </tr>    
    <tr bgcolor="#0066CC">     
      <td colspan="2" height="14">     
        <div align="left"><b>Ҫ��Ա���Ĵ���    
          <select name="play" size="1">    
            <option value="��500��" selected>��500��</option>    
            <option value="��1000��">��1000��</option>    
            <option value="��5000��">��5000��</option>    
          </select>    
          ����ȫ�����ܺ�������</b></div>    
      </td>    
    </tr>    
    <tr>     
      <td colspan="2">     
        <div align="center"><b>     
          ����д���״�ʣ�<br>    
          </b>     
          <textarea name=text cols=50 rows=15></textarea>    
        </div>    
        <p align="center">     
          <input type=submit value=��״ name="submit" style="background-color:FF0000;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">    
          &nbsp;&nbsp;&nbsp;&nbsp; <input type=reset value=��� name="reset" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">    
        </p>    
      </td>    
    </tr>    
  </table>        
</form>        
<div align=center>    
  <p>��</p><hr size=1 width="80%">    
</div>        
</body>        
</html>        
<% end if%>