<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
session("aqjh_sj")=now()
Session("aqjh_jl")=0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 职业 from 用户 where 姓名='"&aqjh_name&"'",conn
if trim(rs("职业"))<>"渔夫" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的职业不是渔夫，所以您不能出海捕渔！！');window.close();</script>"
	response.end
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<html><head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>出海捕鱼</title>
<style type=text/css>
body 
 {
	FONT-SIZE: 12px;CURSOR: url(''); COLOR: #f7deac; SCROLLBAR-FACE-COLOR: #f7deac;SCROLLBAR-ARROW-COLOR: #777777;SCROLLBAR-HIGHLIGHT-COLOR: #555555;SCROLLBAR-SHADOW-COLOR: #555555;SCROLLBAR-3DLIGHT-COLOR: #f7deac;SCROLLBAR-TRACK-COLOR: #333333;SCROLLBAR-DARKSHADOW-COLOR: #f7deac;
}
A:hover {
	COLOR: #ff9900;CURSOR: url('');
}
A.menu:hover {
	COLOR: #ccoooo;CURSOR: url('');
}


td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'f7deac'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff0000;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: 'f7deac'; border: 1 solid 'f7deac'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'f7deac'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'f7deac'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: 'f7deac'; 
	border: f7deacpx solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'f7deac'; color: 'f7deac'; 
	border: 1px solid f7deac;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'f7deac'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'f7deac'}
</style>
</head>

<body bgcolor="#000000" topmargin="0" leftmargin="0">

<table border="0" width="503" height="400" cellspacing="0" cellpadding="0" background="048.gif">
 <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & aqjh_name &"'",conn,3,3
tx=rs("名单头像")
tlj=rs("体力加")
%>
 <tr>
    <td valign="top" height="400">
      <table border="0" width="503" cellspacing="0" cellpadding="0" height="224">
        <tr>
          <td height="224" width="163" valign="top">
            <table border="0" width="165" height="45">
              <tr>
                <td height="41"></td>
              </tr>
            </table>
            <table border="0" width="165" height="102">
              <tr>
                <td width="12" height="98"></td>
                <td width="139" height="98" align="center">&nbsp; <img id=face src="../<%=tx%>" alt=个人形象代表 width='70' height='60'></td> 
              </tr>
            </table>
            <table border="0" width="165" height="57">
              <tr>
                <td height="53" width="19"></td>
                <td height="53" width="132" valign="bottom" align="center"><%=aqjh_name%></td>
              </tr>
            </table>
            　</td>
          <td height="224" width="336" valign="top">
            <table border="0" width="335" height="76">
              <tr>
                <td height="52"></td>
              </tr>
              <tr>
                <td height="16">
                  <div align="center">
                    <center>
                  <table border="1" width="285" bordercolor="#000000" cellspacing="0" cellpadding="0" bordercolordark="#000000" bordercolorlight="#FFFFFF">
                      <tr>
                        <td align="center"><img border="0" src="sailing2.gif" width="240" height="138"></td>
                      </tr>
                    </table>
                    </center>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <div align="center">
        <center>
        <table border="0" width="450" cellspacing="0" height="84">
          <tr>
            <td height="17"></td>
          </tr>
          <tr>
            <td align="center" height="14">※欢迎光临爱情出海捕渔场所※</td>
          </tr>
          <tr>
            <td height="14" align="center"><hr color="#FFFFCC" size="1">
</td>
          </tr>
          <form method=POST action='cwork.asp'>
          <tr>
            <td height="14">
              <table border="0" width="445" cellspacing="0" cellpadding="0">
                <tr>
                  <td><u>出海捕渔可以让你带来不少的财富.还可以偷偷的走私.让你成为巨富.</u></td>
                </tr>
                <tr>
                  <td><u>出海捕渔中可能可以捕到不错的渔喔.有小鲤鱼.大鲨鱼.大草鱼等等</u></td>
                </tr>
                <tr>
                  <td><u>也有可能捕到僵尸牙.红宝石.僵尸血等等非常之多.</u></td>
                </tr>
                <tr>
                  <td><u>但常出海容易生病喔.也有可能让你带来花柳病.所以出海前请三思.</u></td>
                </tr>
              </table>
            </td>
          </tr>
         
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
              <tr>
            <td height="14" align="center">
             
             <table border="0" width="450">
              <tr>
                <td width="224" align="right"><a href="catch1.asp">立刻出海</a></td>
                <td width="224"><a href="javascript:window.close()">算了不出海</td>
              </tr>
             </table>
             
             </td>
          </tr>
             
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="15" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
        </table>
        </center>
      </div>
   
      <p>　</td>
  </tr>
</table>

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
     alert('爱情江湖！欢迎您！本画面不支援滑鼠右键！');                                                                                                                    
     window.event.returnValue=false;                                                                                                                    
     }                                                                                                                    
     document.oncontextmenu=Click;                                                                                            
     </script> 