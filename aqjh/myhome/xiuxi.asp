

<head>
<LINK href="style.css" rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
function show(myshow){Layer2.style.visibility="visible";mess.innerHTML=myshow;}
function hidden(){Layer2.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("欢迎来到爱情江湖！");}}
function shLayers(n,n1){
if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}
if(n1.style.visibility=="visible"){n1.style.visibility="hidden";}
}
</script>
</head>

<body bgcolor="#000000" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="423" height="324" align="left"><img src="IMAGES/fangjian.jpg" width="423" height="324" border="0" usemap="#Map"></td>
</tr>
</table>
<map name="Map">
<area shape="rect" coords="328,8,413,24" href="javascript:shLayers(Layer1,Layer3);" onmouseover="show('如果你是单身侠客,请在这里进行休息,每小时增加体力8万点.在休息的时间里不能够上线操作.时间不限');" onmouseout="hidden();">
<area shape="rect" coords="329,43,413,59" href="javascript:shLayers(Layer3,Layer1);" onmouseover="show('已婚侠客休息的地方,开放时间为21点-23点,时间为:8小时.增加体力100万点.申请生育后由此进行怀孕.');" onmouseout="hidden();">
</map>
<div id="Layer1" style="position: absolute; width: 150px; height: 113px; z-index: 2; left: 61px; top: 156px; background-color: #009900; layer-background-color: #009900; visibility: hidden; background-image: url('IMAGES/bg2.jpg'); layer-background-image: url(IMAGES/bg2.jpg); border: 1px none #000000">
<table width="100%" border="0">
<tr> <form name="form" method="post" action="xiuxiok.asp?act=1">
<td height="106" align="center"><font color="#0000FF">爱情江湖单身侠客休息</font><font color="#FFFFFF"><br>
<br>
<font color="#000000">选择时间: 
<select name="sj">
<option value="1" selected>1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="48">48</option>
<option value="72">72</option>
</select><br>每小时增加体力8万点.</font><br><br><input type="submit" name="Submit" value="确定"></font></td>
</form>
</tr>
</table>
</div>
<div id="Layer2" style="position: absolute; width: 212px; height: 88; z-index: 1; left: 5px; top: 11px; background-image: url('pic/show.jpg'); layer-background-image: url(pic/show.jpg); visibility: hidden; border: 1px none #000000">
<table width="95%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td height="15">&nbsp;</td>
</tr>
<tr>
<td id="mess" height="53">&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>
</table>
</div>
<div id="Layer3" style="position: absolute; width: 150px; height: 113px; z-index: 2; left: 61px; top: 156px; background-color: #FF9900; layer-background-color: #FF9900; visibility: hidden; background-image: url('IMAGES/bg1.jpg'); layer-background-image: url(IMAGES/bg1.jpg); border: 1px none #000000">
<table width="100%" height="115" border="0">
<tr>
<form name="form" method="post" action="xiuxiok.asp?act=2">
<td height="105" align="center"> <font color="#0000FF">爱情江湖已婚侠客休息</font><font color="#FFFFFF"><br>
<br><font color="#000000">21-23点开放,每次休息时间为8小时,增加体力<strong><font color="#0000FF">100</font></strong>万点</font><br><br>
<input type="submit" name="Submit" value="确定"></font></td>
</form>
</tr>
</table>
</div>
