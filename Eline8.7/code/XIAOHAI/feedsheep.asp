<%
Response.Expires=0
if Session("sjjh_name")="" then Response.Redirect "../errors.asp?id=016"
%>
<!--#include file="data1.asp"-->

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
<title>快乐江湖--托儿所♀wWw.happyjh.com♀</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="95%" cellspacing="0" cellpadding="0" height="183" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">　</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="../picc/k3.gif" height="147">　</td>
      <td width="88%" align="center" height="26"><img border="0" src="61.gif" width="111" height="22"></td>
      <td width="4%" background="picc/k3.gif" height="147" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="picc/kg8.jpg" width="520" height="309"></td>
    </tr>
    <tr>
      <td width="88%" height="103" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%" align="center">托儿所</td>                    
          </tr>

          <tr>
            <td width="100%">这儿是喂养小孩的地方，你小孩就住在这儿，看着她无忧无虑的生活着，你觉得在剑侠里不再孤独。不过你每天得来这照顾她哦，不然她会死掉的哦
              
</td>
          </tr>

          <tr>
            <td width="100%" align="center"> 
             <form method="POST" action="checksheep.asp">
         
<table border="0" width="100%">
  <tr>
    <td width="100%">                    <%
rs.open"select sheepname from sheep where user='"&session("sjjh_name")&"'",conn,1,1
if rs.bof then
rs.close
response.write "您还没有领养小孩呢！快领养一个吧"
else
                              %>
                                      <p align="center"><span class="p9">你的孩子的名字叫<font color="#FF0000"><%=rs("sheepname")%> 
                                      </font> 
                                        <input   
                              type="hidden" name="nick" size="20"  
                              style="font-family: Tahoma; font-size: 12px"  
                              value="<%=rs("sheepname")%>">  
                                        </font>  </td>
  </tr>
  <tr>
    <td width="100%"> <p align="center">  
                                        <input type="submit"  
                              value="进入" name="B1"  
                              style="font-family: 宋体; font-size: 12px; background-color: #FFFFCC; border-style: solid; border-width: 1"></td>
  </tr>
</table>
 </form>  
 <%   
rs.close   
end if   
conn.close   
   %>  

</td>
          </tr>

          <tr>
            <td width="100%" align="center">          
   <!--#include file="data1.asp"-->            
                          <p><font color="#FF0000">算工资了：            
                            <%rs.open("select * from sheep where user='"&session("sjjh_name")&"'"),conn,1,1                             
if  rs.bof then%>              
                            <span class="p9">快快养个小孩为你打工赚钱啦！</span>                   
                            <%else                             
 if rs("feedsheepday")>=3then%>              
                            </font></p>              
                          <form method="POST" action="sellmilk.asp">              
                            <p><span class="p9">你孩子为你在孤儿院里打了<%=rs("milk")%>工作单位。你想向孤儿院收多少小时的工钱呢？                
                              <input                 
          type="text" name="milk" size="9"                 
          style="font-family: 宋体; font-size: 12px; background-color: #FFFFCC; border-style: solid; border-width: 1">              
                              <input type="submit"                  
          value="收工钱了" name="B2"                 
          style="font-family: 宋体; font-size: 12px; background-color: #FFFFCC; border-style: solid; border-width: 1">              
                              <input type="hidden"                  
          name="sheepname" value="<%=rs("sheepname")%>">              
                              <%else%>               
                              你的孩子还没找到工作呢</span>。</p>               
                          </form>              
                          <%end if%>              
                                        
                            <%end if%>              
         
</td>
          </tr>

          <tr>
            <td width="100%" align="center">          
</td>
          </tr>
         <td><hr>

          <tr>
            <td width="100%" align="center"></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="88%" height="18"></td>
    </tr>
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">　</td>
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
     alert('快乐江湖欢迎您！');                                                                                                                    
     window.event.returnValue=false;                                                                                                                    
     }                                                                                                                    
     document.oncontextmenu=Click;                                                                                                                    
     </script> 





