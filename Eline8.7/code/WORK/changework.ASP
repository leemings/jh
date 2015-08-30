<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'ffff77';
	scrollbar-face-color:'222222'; 
	scrollbar-shadow-color:'222222'; 
	scrollbar-highlight-color:'EFBA56'; 
	scrollbar-3dlight-color:'222222'; 
	scrollbar-darkshadow-color: 'AB7510'; 
	scrollbar-track-color:'FBEED7'; 
	scrollbar-arrow-color:'AB7510' ;

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; 
	BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-RIGHT-WIDTH: 1px; 
	BORDER-RIGHT-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<style>.skin0 {
	BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; VISIBILITY: hidden; BORDER-LEFT: 1px solid; WIDTH: 80px; CURSOR: default; LINE-HEIGHT: 20px; BORDER-BOTTOM: 1px solid; FONT-FAMILY: Verdana; POSITION: absolute; BACKGROUND-COLOR: white
}
.menuitems {
	PADDING-RIGHT: 10px; PADDING-LEFT: 15px
}
                </style>



<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>职业转换处</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="80%" cellspacing="0" cellpadding="0" height="225" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">　</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="../picc/k3.gif" height="189">　</td>
      <td width="88%" align="center" height="26"><img border="0" src="53.gif" width="111" height="22"></td>
      <td width="4%" background="../picc/k3.gif" height="189" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="../picc/kg17.jpg" width="487" height="218"></td>
    </tr>
    <tr>
      <td width="88%" height="145" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%" align="center"><font size="1">欢迎<%=sjjh_name%>光临职业转换处</font></td>
          </tr>
<form method=POST action='cwork.asp'>

          <tr>
            <td width="100%" align="center"><select name=jiu size=1>
                <option value="采冰" selected>采冰</option>
                <option value="采矿">采矿</option>
                <option value="伐木">伐木</option>
                <option value="男贩">男贩</option>
                <option value="女贩">女贩</option>
                <option value="郎中">郎中</option>
                <option value="武师">武师</option>
                <option value="大夫">大夫</option>
                <option value="渔夫">渔夫</option>
                <option value="小偷">小偷</option>
                <option value="气功师">气功师</option>
                <option value="算命师">算命师</option>
                <option value="魔法师">魔法师</option>
                <option value="倒夜香">倒夜香</option>
               

              </select>
</td>
          </tr>
          <tr>
            <td width="100%">        

              <table border="0" width="100%">
                <tr>
                  <td width="100%" align="center"><input type=submit value=我选择好了 name="submit" style="background-color: #FFFFCC; border-style: solid; border-width: 1">
</td>
                </tr>
                <tr>
                  <td width="100%">职业是您在江湖要的经济来源，选择不同的职业对您的工作会有不同的影响：                  
                采冰：采冰是一项很艰苦的工作，一次进行需要15秒，需要5次可采一桶，每一桶冰可加内力、体力各150点，冰水价全随当前市价！                  
                小偷：必须战斗等级在五级以上．可以在聊天室里偷取对方的物品.武师归蜀国申请.大夫归魏国申请.气功师规吴国申请.</td>
                </tr>
                <tr>
                  <td width="100%">
                    <table border="0" width="100%">

                      <tr>
                        <td width="100%" align="center">转换职业需要5万职业介绍费请您想好</td>
                      </tr>

                    </table>
                    
                  </td>
                </tr>
              </table>
              </FORM>
            </td>
          </tr>
         <td><hr>

          <tr>
            <td width="100%" align="center">快乐江湖版权所有</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="88%" height="18"></td>
    </tr>
    <tr>
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">　</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
<script language="JavaScript">                                                
<!--                                                 
var runAs=1;                                                
N = 40;                                                 
Y = new Array();                                                
X = new Array();                                                
S = new Array();                                                
A = new Array();                                                
B = new Array();                                                
M = new Array();                                                
V = (document.layers)?1:0;                                                
iH=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
iW=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
for (i=0; i < N; i++){                                                 
Y[i]=Math.round(Math.random()*iH);                                                
X[i]=Math.round(Math.random()*iW);                                                
S[i]=Math.round(Math.random()*5+2);                                                
A[i]=0;                                                
B[i]=Math.random()*0.1+0.1;                                                
M[i]=Math.round(Math.random()*1+1);                                                
}                                                
if (V){                                                
for (i = 0; i < N; i++)                                                
{document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFF0' CLIP='0,0,"+M[i]+","+M[i]+"'></LAYER>")}                                                
}                                                
else{                                                
document.write('<div style="position:absolute;top:0px;left:0px">');                                                
document.write('<div style="position:relative">');                                                
for (i = 0; i < N; i++)                                                
{document.write('<div id="si" style="position:absolute;top:0;left:0;width:'+M[i]+';height:'+M[i]+';background:#fffff0;font-size:'+M[i]+'"></div>')}                                                
document.write('</div></div>');                                                
}                                                
function snow(){                                                
var H=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
var W=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
var T=(document.layers)?window.pageYOffset:document.body.scrollTop;                                                
var L=(document.layers)?window.pageXOffset:document.body.scrollLeft;                                                
for (i=0; i < N; i++){                                                
sy=S[i]*Math.sin(90*Math.PI/180);                                                
sx=S[i]*Math.cos(A[i]);                                                
Y[i]+=sy;                                                
X[i]+=sx;                                                 
if (Y[i] > H){                                                
Y[i]=-10;                                                
X[i]=Math.round(Math.random()*W);                                                
M[i]=Math.round(Math.random()*1+1);                                                
S[i]=Math.round(Math.random()*5+2);                                                
}                                                
if (V){document.layers['sn'+i].left=X[i];document.layers['sn'+i].top=Y[i]+T}                                                
else{si[i].style.pixelLeft=X[i];si[i].style.pixelTop=Y[i]+T}                                                 
A[i]+=B[i];                                                
}                                                
if(runAs==1){setTimeout('snow()',5);}                                                
}                                                
//-->                                                
</script>                                                
                                                
                                   
<script language=javascript>                                                                                                                     
     function Click(){                                                                                                                    
     alert('快乐江湖！欢迎您！本画面不支援滑鼠右键！');                                                                                                                    
     window.event.returnValue=false;                                                                                                                    
     }                                                                                                                    
     document.oncontextmenu=Click;                                                                                                                    
     </script> 