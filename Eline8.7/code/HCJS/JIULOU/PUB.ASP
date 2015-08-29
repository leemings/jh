<%@ LANGUAGE=VBScript codepage ="936" %><%


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


%>

<html>

<head>
<link rel="stylesheet" href="../../css.css">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="778">
<tr>
<td bgcolor="#000000" align="right" width="389"><img src="../../images/jd2.gif" width="23" height="107"></td>
<td align="left" colspan="2"><img src="../../images/jd3.gif" width="21" height="107"></td>
</tr>
<tr>
<td bgcolor="#000000" align="right" width="389" height="328" valign="top">
<p align=center style='color:yellow'><font color="#FF9900">欢迎<%=sjjh_name%>光临醉仙楼</font>
<br>
<table border="1" align="center" cellpadding="0" cellspacing="0" width="300" height="200" bordercolor="#008080">
<tr>
<td colspan="3"><font color="#FF9900">这里是京城的名牌酒楼,只见里面喧哗声不断好不热闹.在酒楼门口有一个牌子上面写着:</font></td>
</tr>
<tr align="center">
<td><font color="#FF9900"> 酒 名 </font></td>  
<td><font color="#FF9900"> 体 力 </font></td>
<td><font color="#FF9900"> 银 两 </font></td>
</tr>
<tr align="center">
<td><font color="#FF9900"> 老百干 </font></td>
<td><font color="#FF9900"> 50 </font></td>
<td><font color="#FF9900"> 250 </font></td>
</tr>
<tr align="center">
<td><font color="#FF9900"> 五粮液 </font></td>
<td><font color="#FF9900"> 60 </font></td>
<td><font color="#FF9900"> 300 </font></td>
</tr>
<tr align="center">
<td><font color="#FF9900"> 杜康 </font></td>
<td><font color="#FF9900"> 70 </font></td>
<td><font color="#FF9900"> 350 </font></td>
</tr>
<tr align="center">
<td><font color="#FF9900"> 茅台 </font></td>
<td><font color="#FF9900"> 80 </font></td>
<td bordercolor="#008080"><font color="#FF9900"> 400 </font></td>
</tr>
</table>
</td>
<td align="left" height="328"></td>
<td align="center" width=389 height="328" valign="top">
<a href="http://www.jdsky.net"><img border="0" src="../../chat/img/110.jpg"></a>  <form method=POST action='pub1.asp'>
<table width="300">
<tr>
<td>
<tr>
<td align=center><font color="#FF0000">你向小二买下了：</font> 
<select name=jiu size=1>
<option value="lao">老百干
<option value="wu">五粮液
<option value="du">杜康
<option value="mao">茅台
</select>
  <font color="#FF0000">就举起酒杯</font></td>
</tr>
<tr>
<td  align=center>
<input type=submit value=一饮而尽 name="submit">
</td>
</tr>
<tr>
<td > <font color="#008080"> 1、在酒楼喝酒可以增加体力； <br>
2、酒虽然是个好东西,但是不能贪多； 酒量不够请少饮,否则,嘻嘻......<img border="0" src="../../chat/pic/28.gif"><br>
3、酒量与你的武功和经验有关！ 酒量=(武功+经验)/100 要想不醉倒须 酒量>1，自己掂量掂量吧！</font></td> 
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>


