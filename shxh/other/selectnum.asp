<html>
<head>
<link rel="stylesheet" href="../style.css">
<script language=javascript>
function rnd(){
	for(var i=0;i<30;i++){
		document.form1.elements[i].value=""+Math.round(Math.random()*9);
	}
}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>'>
<form action='selectnum2.asp' method=post name=form1>
<table width=100% border=3>
<tr align=center bgcolor=ffff00><td>编号</td><td>号码</td></tr>
<tr><td>1</td><td><input type=text size=1 maxlength=1 name='num00' value=''><input type=text size=1 maxlength=1 name='num01' value=''><input type=text size=1 maxlength=1 name='num02' value=''><input type=text size=1 maxlength=1 name='num03' value=''><input type=text size=1 maxlength=1 name='num04' value=''><input type=text size=1 maxlength=1 name='num05' value=''></td></tr>
<tr><td>2</td><td><input type=text size=1 maxlength=1 name='num10' value=''><input type=text size=1 maxlength=1 name='num11' value=''><input type=text size=1 maxlength=1 name='num12' value=''><input type=text size=1 maxlength=1 name='num13' value=''><input type=text size=1 maxlength=1 name='num14' value=''><input type=text size=1 maxlength=1 name='num15' value=''></td></tr>
<tr><td>3</td><td><input type=text size=1 maxlength=1 name='num20' value=''><input type=text size=1 maxlength=1 name='num21' value=''><input type=text size=1 maxlength=1 name='num22' value=''><input type=text size=1 maxlength=1 name='num23' value=''><input type=text size=1 maxlength=1 name='num24' value=''><input type=text size=1 maxlength=1 name='num25' value=''></td></tr>
<tr><td>4</td><td><input type=text size=1 maxlength=1 name='num30' value=''><input type=text size=1 maxlength=1 name='num31' value=''><input type=text size=1 maxlength=1 name='num32' value=''><input type=text size=1 maxlength=1 name='num33' value=''><input type=text size=1 maxlength=1 name='num34' value=''><input type=text size=1 maxlength=1 name='num35' value=''></td></tr>
<tr><td>5</td><td><input type=text size=1 maxlength=1 name='num40' value=''><input type=text size=1 maxlength=1 name='num41' value=''><input type=text size=1 maxlength=1 name='num42' value=''><input type=text size=1 maxlength=1 name='num43' value=''><input type=text size=1 maxlength=1 name='num44' value=''><input type=text size=1 maxlength=1 name='num45' value=''></td></tr>
<tr><td colspan=2 align=center><input type=button value='随机号码' onclick='javascript:rnd();'> <input type=submit value='我要发财了'> <input type=reset value='重选'></td></tr>
</table>
</form>
<p align=center><a href="javascript:location.replace('lottery.asp');">返回</a></p>
</body>
</body>