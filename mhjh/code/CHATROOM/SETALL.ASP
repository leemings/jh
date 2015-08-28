<html>
<head>
<link rel="stylesheet" href="css.css">
<title></title>
</head>
<body oncontextmenu='self.event.returnValue=false'  leftmargin="5" topmargin="7" marginwidth="5" marginheight="5" background='bg1.gif'>
<div align="center">
  <center>
  <table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%"><form name=talkform action="sendmsg.asp" method=post target="sendfrm" onsubmit="javascript:return(checklingual());">
 <br> <br>选择背景色<br>
<select name="bgcolor" onchange="parent.chgbgcolor(this.value);">                                                 
<option value="#FFFFFF" style='background:#FFFFFF;color=#FFFFFF'>背景色1</option>                                                 
<option value="#C8C8FF" style='background:#C8C8FF;color=#C8C8FF'>背景色2</option>                                                 
<option value="#FFFFCC" style='background:#FFFFCC;color=#FFFFCC'>背景色3</option>                                                 
<option value="#F4F4FF" style='background:#F4F4FF;color=#F4F4FF'>背景色4</option>                                                 
<option value="#33FFFF" style='background:#33FFFF;color=#33FFFF'>背景色5</option>                                                 
<option value="#FFF4F4" style='background:#FFF4F4;color=#FFF4F4'>背景色6</option>                                                 
<option value="#F4FFF4" style='background:#F4FFF4;color=#F4FFF4'>背景色7</option>                                                 
<option value="#F4F4FF" style='background:#F4F4FF;color=#F4F4FF'>背景色8</option>                                                 
<option value="#EEEEEE" style='background:#EEEEEE;color=#EEEEEE'>背景色9</option>                                                 
</select><br>
 <p><br>
选择屏幕特效<br>    
	<select name=refresh1 onchange='parent.refresh1(this.value);' style='font-size:12px'>    
    <option value="" selected>屏幕特效    
    <option value="STARMOON.js">星月传说    
    <option value="js.js">我心飞扬    
	<option value="XMAS.js">新年快乐    
	<option value="test.js">快乐天使    
    <option value="snow.js">樱花纷飞    
    <option value="ll.js">彩云飘飘    
    </select><br></p>
 <p><br>
    选择屏幕字体<br>
    	<select name=ziti onchange='parent.ziti(this.value);' style='font-size:12px' size="1">    
    <option selected>请选择</option>
    <option value="css.css">小字体    
    <option value="Css2.css">大字体         
    </select>  
 </p>
</form></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
