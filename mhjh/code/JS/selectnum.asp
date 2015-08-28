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
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<form action='selectnum2.asp' method=post name=form1>
<p align="center">&nbsp;<img border="0" src="../IMAGE/252.gif">&nbsp; <img border="0" src="../IMAGE/loc54.gif">&nbsp;&nbsp;&nbsp;&nbsp; 
<img border="0" src="../IMAGE/251.GIF"></p> 
<div align="center"> 
<center> 
<table width=405 border=3> 
<tr align=center bgcolor=ffff00><td bordercolor="#FFFFFF" bgcolor="#00CC66" width="59"><font color="#FFFF00">编号</font></td><td bordercolor="#FFFFFF" bgcolor="#00CC66" width="326"><font color="#FFFF00">号码</font></td></tr> 
<tr><td bgcolor="#005B00" width="59">1</td><td bgcolor="#005B00" width="326"><input type=text size=1 maxlength=1 name='num00' value=''><input type=text size=1 maxlength=1 name='num01' value=''><input type=text size=1 maxlength=1 name='num02' value=''><input type=text size=1 maxlength=1 name='num03' value=''><input type=text size=1 maxlength=1 name='num04' value=''><input type=text size=1 maxlength=1 name='num05' value=''></td></tr> 
<tr><td bgcolor="#005B00" width="59">2</td><td bgcolor="#005B00" width="326"><input type=text size=1 maxlength=1 name='num10' value=''><input type=text size=1 maxlength=1 name='num11' value=''><input type=text size=1 maxlength=1 name='num12' value=''><input type=text size=1 maxlength=1 name='num13' value=''><input type=text size=1 maxlength=1 name='num14' value=''><input type=text size=1 maxlength=1 name='num15' value=''></td></tr> 
<tr><td bgcolor="#005B00" width="59">3</td><td bgcolor="#005B00" width="326"><input type=text size=1 maxlength=1 name='num20' value=''><input type=text size=1 maxlength=1 name='num21' value=''><input type=text size=1 maxlength=1 name='num22' value=''><input type=text size=1 maxlength=1 name='num23' value=''><input type=text size=1 maxlength=1 name='num24' value=''><input type=text size=1 maxlength=1 name='num25' value=''></td></tr> 
<tr><td bgcolor="#005B00" width="59">4</td><td bgcolor="#005B00" width="326"><input type=text size=1 maxlength=1 name='num30' value=''><input type=text size=1 maxlength=1 name='num31' value=''><input type=text size=1 maxlength=1 name='num32' value=''><input type=text size=1 maxlength=1 name='num33' value=''><input type=text size=1 maxlength=1 name='num34' value=''><input type=text size=1 maxlength=1 name='num35' value=''></td></tr> 
<tr><td bgcolor="#005B00" width="59">5</td><td bgcolor="#005B00" width="326"><input type=text size=1 maxlength=1 name='num40' value=''><input type=text size=1 maxlength=1 name='num41' value=''><input type=text size=1 maxlength=1 name='num42' value=''><input type=text size=1 maxlength=1 name='num43' value=''><input type=text size=1 maxlength=1 name='num44' value=''><input type=text size=1 maxlength=1 name='num45' value=''></td></tr> 
<tr><td colspan=2 align=center bgcolor="#005B00" width="391"><input type=button value='随机号码' onclick='javascript:rnd();'> <input type=submit value='我要发财了'> <input type=reset value='重选'></td></tr> 
</table> 
</center> 
</div> 
<p align="center">你可以自己填写心目中的号码或者由系统随机生成号码，祝你好运！</p> 
</form> 
<p align=center><a href="javascript:location.replace('lottery.asp');">返回</a></p> 
<p align="center">　</p> 
</body> 
