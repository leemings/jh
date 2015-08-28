<html>
<head>
<BGSOUND src="mid/midi1.mid" loop=-1>
<title>魔幻江湖</title>
<STYLE type=text/css>
body {MARGIN: 0px; font-family: "宋体", "serif"; font-size:9; line-height:100%;}
TD {FONT-SIZE: 9pt}
A:link {FONT-SIZE: 9pt; COLOR: #999966; TEXT-DECORATION: none}
A:visited {FONT-SIZE: 9pt; COLOR: #999966; TEXT-DECORATION: none}
A:hover {FONT-SIZE: 9pt; COLOR: #990000; TEXT-DECORATION: underline}
.login {BORDER-RIGHT: 0px; BORDER-TOP: 0px; FONT-SIZE: 9pt; BACKGROUND: #000000; BORDER-LEFT: 0px; COLOR: #cccccc; BORDER-BOTTOM: #cccccc 1px solid}
.submit {BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; FONT-SIZE: 9pt; BACKGROUND: #000000; BORDER-LEFT: #cccccc 1px solid; WIDTH: 50px; COLOR: #cccccc; BORDER-BOTTOM: #cccccc 1px solid; HEIGHT: 20px}
</STYLE>
<SCRIPT language=JScript>
<!--
function check(){
if(document.forms[0].username.value.length<=1){alert('请填入您已经申请成功的用户名称！');document.forms[0].username.select();return false;}
else if(document.forms[0].username.value.indexOf("\'")!=-1||document.forms[0].username.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].username.select();return false;}
else if(document.forms[0].password.value.length<6){alert('密码过短，请验证是否输入正确');document.forms[0].password.select();return false;}
else if(document.forms[0].password.value.indexOf("\'")!=-1||document.forms[0].password.value.indexOf("\"")!=-1){alert('请勿使用非法字符，好吗？');document.forms[0].password.select();return false;}
else {return true;}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

//-->
</SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- Fireworks MX Dreamweaver MX target.  Created Fri Jul 25 01:06:41 GMT+0800 (?D1ú±ê×?ê±??) 2003-->
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="10" topmargin="0">
<div id="Layer1" style="position:absolute; width:189; height:222; z-index:1; left: 64; top: 27"> 
  <table width="195"  border="2" height="218" bordercolor="#00FFFF" style="border:1px dotted #008000; border-collapse: collapse" cellpadding="0" cellspacing="0">
     <tr> 
      <form action="login.asp" method=post onsubmit='return(check());' name=form1>
        <input type=hidden name=reg1 value="1236">
        <TD height="206" width="193" align="center" valign="top" style="border: 1px dotted #800080">
          <script language=JScript src="onlinelist.asp"></script><br>
          &nbsp;<FONT color=#FFFF00>姓 名：</FONT><FONT color=#ffffff> 
          <INPUT type="text" name="username"  maxlength="10"                                                                                         
      style="color: #000000; background-color: #FFFFFF; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'" size="10">
          <br>&nbsp;<FONT color=#FFFF00>密 码：</FONT></FONT> <FONT color=#ffffff> 
          <INPUT name="password" type="password"                                                                                                                       
      style="color: #000000; background-color: #FFFFFF; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'" size="10" maxlength="15" >
          <br><font color="#FFFF00">&nbsp;认 证：</font>
          <input name=reg type=text style="color: #000000; background-color: #FFFFFF; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#FFFFFF'" value="1236" size=10 maxlength="4" > 
          <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT class=submit type=submit value="登 录" name=Submit> 
      <p align="left">　<p align="left">　<p align="left" style="line-height: 150%">&nbsp; &nbsp;</font>版&nbsp; 权：<FONT color=#ffffff><font color="#000000"><a target="_blank" href="http://www.yx8.net"><font color="#FF0000">游戏吧</font></a> 
      </font><br>&nbsp;</font><font color="#FFFFFF">&nbsp; </font>授权给：<FONT color=#ffffff><font color="#000000"><% response.write Application("yx8_mhjh_systemname")%> 
      </font><br>&nbsp; &nbsp;</font>站&nbsp; 长：<FONT color=#ffffff><font color="#000000"><%=Application("yx8_mhjh_admin")%> 
      </font><br>&nbsp;&nbsp; </font>序列号：<FONT color=#ffffff><font color="#000000"><% response.write Application("yx8_mhjh_seriesnumber")%></font></TD>
      </FORM>
    </tr>
    </table>
</div>
<div id="Layer2" style="position:absolute; width:149; height:45; z-index:2; left: 84; top: 146"> 
  <table width="145"  border="0" height="43">
    <tr> 
      <td width="57" height="8"><a href="#" onClick="MM_openBrWindow('gr/regstep1.asp','reg','width=640,height=520')">
      <font color="#FFFF00"><strong style="font-weight: 400">注册帐号</strong></font></a></td>
      <td width="78" height="8"><a href="#" onClick="MM_openBrWindow('gr/help.asp','help','width=640,height=470')">
      <font color="#FFFF00"><strong style="font-weight: 400">掉线自救</strong></font></a></td>
    </tr>
    <tr> 
      <td width="57" height="15"><a href="#" onClick="MM_openBrWindow('gr/chgpw.asp','pw','width=640,height=450')">
      <font color="#FFFF00"><strong style="font-weight: 400">修改密码</strong></font></a></td>
      <td width="78" height="15"><a href="#" onClick="MM_openBrWindow('gr/relive.asp','relive','width=640,height=440')">
      <font color="#FFFF00"><strong style="font-weight: 400">战将复活</strong></font></a></td>
    </tr>
    <tr> 
      <td width="57" height="8"><a href="#" onClick="MM_openBrWindow('gr/getzh.asp','relive','width=640,height=450')">
      <font color="#FFFF00"><strong style="font-weight: 400">查找帐号</strong></font></a></td>
      <td width="78" height="8"><a href="#" onClick="MM_openBrWindow('main.asp','relive','width=580,height=490')">
      <font color="#FFFF00"><strong style="font-weight: 400">参观江湖</strong></font></a></td>
    </tr>
  </table>
</div>
<table border="0" cellpadding="0" cellspacing="0" width="776">
<!-- fwtable fwsrc="未命名" fwbase="50.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  
  <tr>
   <td><img name="n50_r1_c1" src="images/50_r1_c1.jpg" width="106" height="70" border="0"></td>
   <td><img name="n50_r1_c2" src="images/50_r1_c2.jpg" width="94" height="70" border="0"></td>
   <td><img name="n50_r1_c3" src="images/50_r1_c3.jpg" width="104" height="70" border="0"></td>
   <td><img name="n50_r1_c4" src="images/50_r1_c4.jpg" width="98" height="70" border="0"></td>
   <td><img name="n50_r1_c5" src="images/50_r1_c5.jpg" width="98" height="70" border="0"></td>
   <td><img name="n50_r1_c6" src="images/50_r1_c6.jpg" width="100" height="70" border="0"></td>
   <td><img name="n50_r1_c7" src="images/50_r1_c7.jpg" width="100" height="70" border="0"></td>
   <td><img name="n50_r1_c8" src="images/50_r1_c8.jpg" width="76" height="70" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r2_c1" src="images/50_r2_c1.jpg" width="106" height="72" border="0"></td>
   <td><img name="n50_r2_c2" src="images/50_r2_c2.jpg" width="94" height="72" border="0"></td>
   <td><img name="n50_r2_c3" src="images/50_r2_c3.jpg" width="104" height="72" border="0"></td>
   <td><img name="n50_r2_c4" src="images/50_r2_c4.jpg" width="98" height="72" border="0"></td>
   <td><img name="n50_r2_c5" src="images/50_r2_c5.jpg" width="98" height="72" border="0"></td>
   <td><img name="n50_r2_c6" src="images/50_r2_c6.jpg" width="100" height="72" border="0"></td>
   <td><img name="n50_r2_c7" src="images/50_r2_c7.jpg" width="100" height="72" border="0"></td>
   <td><img name="n50_r2_c8" src="images/50_r2_c8.jpg" width="76" height="72" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r3_c1" src="images/50_r3_c1.jpg" width="106" height="74" border="0"></td>
   <td valign="top"><img name="n50_r3_c2" src="images/50_r3_c2.jpg" width="94" height="74" border="0"></td>
   <td><img name="n50_r3_c3" src="images/50_r3_c3.jpg" width="104" height="74" border="0"></td>
   <td><img name="n50_r3_c4" src="images/50_r3_c4.jpg" width="98" height="74" border="0"></td>
   <td><img name="n50_r3_c5" src="images/50_r3_c5.jpg" width="98" height="74" border="0"></td>
   <td><img name="n50_r3_c6" src="images/50_r3_c6.jpg" width="100" height="74" border="0"></td>
   <td><img name="n50_r3_c7" src="images/50_r3_c7.jpg" width="100" height="74" border="0"></td>
   <td><img name="n50_r3_c8" src="images/50_r3_c8.jpg" width="76" height="74" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r4_c1" src="images/50_r4_c1.jpg" width="106" height="54" border="0"></td>
   <td><img name="n50_r4_c2" src="images/50_r4_c2.jpg" width="94" height="54" border="0"></td>
   <td><img name="n50_r4_c3" src="images/50_r4_c3.jpg" width="104" height="54" border="0"></td>
   <td><img name="n50_r4_c4" src="images/50_r4_c4.jpg" width="98" height="54" border="0"></td>
   <td><img name="n50_r4_c5" src="images/50_r4_c5.jpg" width="98" height="54" border="0"></td>
   <td><img name="n50_r4_c6" src="images/50_r4_c6.jpg" width="100" height="54" border="0"></td>
   <td><img name="n50_r4_c7" src="images/50_r4_c7.jpg" width="100" height="54" border="0"></td>
   <td><img name="n50_r4_c8" src="images/50_r4_c8.jpg" width="76" height="54" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r5_c1" src="images/50_r5_c1.jpg" width="106" height="72" border="0"></td>
   <td><img name="n50_r5_c2" src="images/50_r5_c2.jpg" width="94" height="72" border="0"></td>
   <td><img name="n50_r5_c3" src="images/50_r5_c3.jpg" width="104" height="72" border="0"></td>
   <td><img name="n50_r5_c4" src="images/50_r5_c4.jpg" width="98" height="72" border="0"></td>
   <td><img name="n50_r5_c5" src="images/50_r5_c5.jpg" width="98" height="72" border="0"></td>
   <td><img name="n50_r5_c6" src="images/50_r5_c6.jpg" width="100" height="72" border="0"></td>
   <td><img name="n50_r5_c7" src="images/50_r5_c7.jpg" width="100" height="72" border="0"></td>
   <td><img name="n50_r5_c8" src="images/50_r5_c8.jpg" width="76" height="72" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r6_c1" src="images/50_r6_c1.jpg" width="106" height="80" border="0"></td>
   <td><img name="n50_r6_c2" src="images/50_r6_c2.jpg" width="94" height="80" border="0"></td>
   <td><img name="n50_r6_c3" src="images/50_r6_c3.jpg" width="104" height="80" border="0"></td>
   <td><img name="n50_r6_c4" src="images/50_r6_c4.jpg" width="98" height="80" border="0"></td>
   <td><img name="n50_r6_c5" src="images/50_r6_c5.jpg" width="98" height="80" border="0"></td>
   <td><img name="n50_r6_c6" src="images/50_r6_c6.jpg" width="100" height="80" border="0"></td>
   <td><img name="n50_r6_c7" src="images/50_r6_c7.jpg" width="100" height="80" border="0"></td>
   <td><img name="n50_r6_c8" src="images/50_r6_c8.jpg" width="76" height="80" border="0"></td>
  </tr>
  <tr>
   <td><img name="n50_r7_c1" src="images/50_r7_c1.jpg" width="106" height="86" border="0"></td>
   <td><img name="n50_r7_c2" src="images/50_r7_c2.jpg" width="94" height="86" border="0"></td>
   <td><img name="n50_r7_c3" src="images/50_r7_c3.jpg" width="104" height="86" border="0"></td>
   <td><img name="n50_r7_c4" src="images/50_r7_c4.jpg" width="98" height="86" border="0"></td>
   <td><img name="n50_r7_c5" src="images/50_r7_c5.jpg" width="98" height="86" border="0"></td>
   <td><img name="n50_r7_c6" src="images/50_r7_c6.jpg" width="100" height="86" border="0"></td>
   <td><img name="n50_r7_c7" src="images/50_r7_c7.jpg" width="100" height="86" border="0"></td>
   <td><img name="n50_r7_c8" src="images/50_r7_c8.jpg" width="76" height="86" border="0"></td>
  </tr>
</table>
</body>
</html>