<%
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../../error.asp?id=016"
if usercorp<>"官府" then Response.Redirect "../../exit.asp"
%>
<html>
<head>
<title>问题上转</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="home.css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_validateForm() { //v3.0
var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
if (val) { nm=val.name; if ((val=val.value)!="") {
if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
} else if (test!='R') { num = parseFloat(val);
if (val!=''+num) errors+='- '+nm+' must contain a number.\n';
if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
min=test.substring(8,p); max=test.substring(p+1);
if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
} } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
} if (errors) alert('The following error(s) occurred:\n'+errors);
document.MM_returnValue = (errors == '');
}
//-->
</script>
<LINK href="../../style.css" rel=stylesheet>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<table width="468" border="0" cellspacing="0" cellpadding="0" align="center" height="300">
<tr valign="top">
<td>
<h3>问题上传</h3>
<form name="form" action="updata.asp?action=add" method="post" >
<table width="100%" border="0" cellspacing="0" cellpadding="2">
<tr>
<td width="5%" nowrap>问题类型</td>
<td width="95%">
<input type="radio" name="lx" value="1" <%if session("lx")<="1" then response.write "checked"%> >
1艺术  
<input type="radio" name="lx" value="2"<%if session("lx")="2" then response.write "checked"%>>  
2体育  
<input type="radio" name="lx" value="3"<%if session("lx")="3" then response.write "checked"%>>  
3文学  
<input type="radio" name="lx" value="4"<%if session("lx")="4" then response.write "checked"%>>  
4科学  
<input type="radio" name="lx" value="5"<%if session("lx")="5" then response.write "checked"%>>  
5综合  
<input type="radio" name="lx" value="6"<%if session("lx")="6" then response.write "checked"%>>  
6游戏动漫</td>  
</tr>  
<tr>  
<td width="5%" nowrap>题目</td>  
<td width="95%">  
<input type="text" name="question" size="50" maxlength="100">  
</td>  
</tr>  
<tr>  
<td width="5%" nowrap>答案A</td>  
<td width="95%">  
<input type="text" size="50" maxlength="100" name="A">  
</td>  
</tr>  
<tr>  
<td width="5%" nowrap>答案B</td>  
<td width="95%">  
<input type="text" name="B" size="50" maxlength="100">  
</td>  
</tr>  
<tr>  
<td width="5%" nowrap>答案C</td>  
<td width="95%">  
<input type="text" name="C" size="50" maxlength="100">  
</td>  
</tr>  
<tr>  
<td width="5%" nowrap>答案D</td>  
<td width="95%">  
<input type="text" name="D" size="50" maxlength="100">  
</td>  
</tr>  
<tr>  
<td width="5%" nowrap>正确答案</td>  
<td width="95%">  
<input type="radio" name="ANSWER" value="A" checked>  
A  
<input type="radio" name="ANSWER" value="B">  
B  
<input type="radio" name="ANSWER" value="C">  
C  
<input type="radio" name="ANSWER" value="D">  
D </td>  
</tr>  
<tr>  
<td width="5%" nowrap>&nbsp;</td>  
<td width="95%">  
<input type="submit" name="Submit" value="上传" onClick="MM_validateForm('question','','R','A','','R','B','','R','C','','R','D','','R');return document.MM_returnValue">  
</td>  
</tr>  
</table>  
</form>  
<h3>&nbsp;</h3>  
</td> 
</tr> 
</table> 
</body> 
</html> 
