<%
Response.Expires=0
if Session("sjjh_name")="" then Response.Redirect "../../errors.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sqlstr="SELECT * FROM �û� where ����='"&Session("sjjh_name")&"'"
Set Rs=conn.Execute(sqlstr)
peiou=rs("��ż")
if peiou="��" then
Response.Write "<script language=javascript>alert('�㻹û����ż��������������');history.back();</script>"
else
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
<title>�����¶�Ժ��wWw.happyjh.com��</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="95%" cellspacing="0" cellpadding="0" height="225" background="picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">��</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="picc/k3.gif" height="189">��</td>
      <td width="88%" align="center" height="26"><img border="0" src="61.gif" width="111" height="22"></td>
      <td width="4%" background="picc/k3.gif" height="189" rowspan="4">��</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="picc/kg8.jpg" width="520" height="309"></td>
    </tr>
    <tr>
      <td width="88%" height="145" valign="top">
        <table border="0" width="100%">

          <tr>
              <td width="100%"><font color="#FFFFFF" size="2">����ǽ��������Ĺ¶�Ժ�����ڴ��ڽ������������Թ¶�Ժ��Ĺ¶��ܶ࣬������������10����������һ���¶��ؼҡ���ÿ��ι�������չ�������ĺ��ӳ���ᱨ����ġ��ȵ�һ��ʱ�������򹤸���׬��ҵġ�</font></td>                        
          </tr> 
 
          <tr> 
            <td width="100%" align="center">�����¶�Ժ��Ĺ¶�</td> 
          </tr> 
 
          <tr> 
            <td width="100%">�����¶�:Ϊ��ĺ���ȡһ�����������ְɣ�<form name="form1" method="POST" action="buysheep.asp">                       
                                            <input type="text" name="sheepname" size="25"                       
                        style="font-family: ����; font-size: 12px; background-color: #FFFFCC; border-style: solid; border-width: 1">                       
                                            <input                        
                        type="button" value="����" name="B1"                       
                        style="font-family: ����; font-size: 12px; background-color: #FFFFCC; border-style: solid; border-width: 1">                       
                                          </form>                       
                                          <script language="vbscript">                        
<!--                        
sub b1_onclick                        
if form1.sheepname.value="" then                        
msgbox"С�������ֲ���Ϊ��"                        
else                        
form1.submit                        
end if                        
end sub                        
-->                        
                                          </script>                       
</td> 
          </tr>  <!--#include file="data1.asp"-->   
          <tr> 
            <td width="100%" align="center">  
 
                     
                          <p><font color="#FFCC00"><span class="p9">�㹤���ˣ�</span>                        
                            <%rs.open("select * from sheep where user='"&session("sjjh_name")&"'"),conn,1,1                                          
if  rs.bof then%>           
rs.open "select * from �û� where ����='" & to1 &"'" &" and ����='" & mp & "'",conn,2,2  
               
                            <span class="p9">�������С��Ϊ��׬Ǯ����</span>                               
                            <%else                                          
 if rs("feedsheepday")>=3then%>                         
                            </font></p>                         
                          <form method="POST" action="sellmilk.asp">                         
                            <p><span class="p9">�㺢��Ϊ���ڹ¶�Ժ�����<%=rs("milk")%>������λ��������¶�Ժ�ն���Сʱ�Ĺ�Ǯ�أ�                          
                              <input                            
          type="text" name="milk" size="9"                            
          style="font-family: ����; font-size: 12px">                         
                              <input type="submit"                            
          value="�չ�Ǯ��" name="B2"                            
          style="font-family: ����; font-size: 12px">                         
                              <input type="hidden"                            
          name="sheepname" value="<%=rs("sheepname")%>">                         
                              <%else%>                         
                              ��ĺ��ӻ�û�ҵ�������</span>��</p>                         
                          </form>                         
                          <%end if%>                         
                                                   
                            <%end if%>                         
 
 
 
 
          <tr> 
            <td width="100%" align="center"></td> 
          </tr> 
          <tr> 
            <td width="100%">         
 
              </FORM> 
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
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">��</td> 
    </tr> 
  </table> 
  </center> 
</div> 
 
</body> 
 
</html> 
 <%        
end if        
rs.Close        
set rs=nothing        
conn.Close        
set conn=nothing        
%> 
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
     alert('���ֽ�����ӭ����');                                                                                                                     
     window.event.returnValue=false;                                                                                                                     
     }                                                                                                                     
     document.oncontextmenu=Click;                                                                                                                     
     </script>  
 
 
 
 
 
 
 
 
 
 
 
