<!--#include file="session.asp"-->
<html>
<head>
<title>�����б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="style.css" type="text/css">
<SCRIPT language=javascript>
var intLastId = '0'
function doClick(intType) {
	var i, srcElement, targetElement;
	srcElement = window.event.srcElement;
	if (intType == 1) i = srcElement.id;
	if (intType == 2) i = srcElement.id.substring(7);
	if (intLastId !='0' && intLastId != i && !(intLastId.charAt(0) == i.charAt(0)) ) {
		document.all(intLastId).src = "images/plus1.gif";
		document.all("subject" + intLastId + "_message").style.display = "none";
	}
	intLastId = i;

	targetElement = document.all("subject" + i + "_message");
	if (targetElement.style.display == "none") {
		document.all(i).src = "images/minus1.gif";
		targetElement.style.display = "";	
	} 
	else {
		document.all(i).src = "images/plus1.gif";
		targetElement.style.display = "none";
	}
}
</SCRIPT>
</head>

<body bgcolor="#D0DCE0" text="#000000" leftmargin="5" topmargin="5">
<table width="130" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="23"><b>&gt;&gt; <a href="welcome.asp" target="right">����˵�</a></b></td>
</tr>
  <tr>
<td>
<!-- ���ֹ������ -->
<!-- ��������ʼ -->
<img id=4 style="CURSOR: hand" onClick=doClick(1); src="images/plus1.gif" width="31" height="15" align="absmiddle">
<span id=subject4 style="CURSOR: hand; color:#000000" onclick=doClick(2);> ����Ƶ������</SPAN><br> 
      <div id=subject4_message style="DISPLAY: none"> <img src="images/minus2.gif" width="50" height="16" align="absmiddle">  
        <a href="manage/add_flash_class.asp" target="right">������</a><br>
        <img src="images/minus2.gif" width="50" height="16" align="absmiddle"> 
        <a href="manage/manage_flash_class.asp" target="right">������</a><br>
        <img src="images/minus2.gif" width="50" height="16" align="absmiddle"> 
        <a href="manage/add_flash.asp" target="right">��Ӷ���</a><br>
        <img src="images/minus2.gif" width="50" height="16" align="absmiddle"> 
        <a href="manage/manage_flash.asp" target="right">��������</a> </div>
<!-- ��������� -->
<!-- ϵͳ����ʼ -->
<img id=18 style="CURSOR: hand" onClick=doClick(1); src="images/plus1.gif" width="31" height="15" align="absmiddle">
<span id=subject18 style="CURSOR: hand; color:#0000ff" onclick=doClick(2);> ϵͳ�ʺŹ���</SPAN> 
      <div id=subject18_message style="DISPLAY: none"> <img src="images/minus2.gif" width="50" height="16" align="absmiddle">  
        <a href="add_user.asp" target="right">����ʺ�</a><br> <img src="images/minus2.gif" width="50" height="16" align="absmiddle">  
        <a href="user.asp" target="right">�����ʺ�</a> 
 </div> 
<!-- ϵͳ������� --> 
<img src="images/MINUS1.GIF" width="31" height="15" align="absmiddle"> <a href="personal.asp" target="right"><font color="blue">���������޸�</font></a> 
<img src="images/MINUS1.GIF" width="31" height="15" align="absmiddle"> <a href="/" target="_parent"><font color="blue">������վ��ҳ</font></a> 
<img src="images/MINUS1.GIF" width="31" height="15" align="absmiddle"> <a href="logout.asp" target="_parent"><font color="red">�˳�����ϵͳ</font></a></td> 
</tr> 
</table> 
</html>