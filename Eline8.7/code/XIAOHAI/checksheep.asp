<!--#include file="data2.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("nick"),"С������") 
%>
<%
if session("sjjh_name")="" then
response.redirect"warning.htm"
else

IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request.form("nick")
if sheepname="" or len(sheepname)="" then
%>
<script language="Vbscript">
msgbox"С����������д����ȷ��",0,"��ʾ"
history.back
</script>
<%
else
rs.open"select * from rules",conn,1,1
if rs.bof then
rs.close
conn.close
response.write"û�й������"
else
dellifeday=rs("dellifeday")
rs.close
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'",conn,1,1
if rs.bof then
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"���������С�������ˣ�",0,"��ʾ"
history.back
</script>
<%
else
logintoday=rs("logintoday")
feeddate=rs("feeddate")
life=rs("life")
hungry=rs("hungry")
clean=rs("clean")
sheephappy=rs("sheephappy")
sheephealth=rs("sheephealth")
rs.close
if datediff("d",logintoday,date)<>0 then
tempdatetoday=date
conn.execute"update sheep set logintoday='"&tempdatetoday&"' where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"
temptime=datediff("d",feeddate,date)
templife=life-dellifeday*temptime
temphungry=hungry-dellifeday*temptime/2
if temphungry<0 then
temphungry=0
end if
tempclean=clean-dellifeday*temptime/2
if tempclean<0 then
tempclean=0
end if
tempsheephappy=sheephappy-dellifeday*temptime/2
if tempsheephappy<0 then
tempsheephappy=0
end if
tempsheephealth=sheephealth-dellifeday*temptime/2
if tempsheephealth<0 then
tempsheephealth=0
end if
'conn.execute"update sheep set clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"',hungry='"&temphungry&"' where sheepname='"&sheepname&"' and id="&tempid&" and user='"&session("sjjh_name")&"'"
conn.execute"update sheep set life='"&templife&"',hungry='"&temphungry&"',clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"' where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"
end if
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"
if rs("life")<=0 or rs("clean")<=0 or rs("sheephappy")<=0 or rs("sheephealth")<=0 or rs("hungry")<=0 then
rs.close
conn.execute"delete * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"
conn.close
%>
<script language="vbscript">
msgbox"������û�о����չ�С��������С���Ѿ����ˣ���������һ���ɡ�",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'",conn,1,1
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
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
<title>�����¶�</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="95%" cellspacing="0" cellpadding="0" height="183" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">��</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="picc/k3.gif" height="147">��</td>
      <td width="88%" align="center" height="26"><img border="0" src="61.gif" width="111" height="22"></td>
      <td width="4%" background="picc/k3.gif" height="147" rowspan="4">��</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="picc/kg8.jpg" width="564" height="309"></td>
    </tr>
    <tr>
      <td width="88%" height="103" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="25%">��-С������ֵ��</td>
                  <td width="25%"> <%
if rs("life")>80 then
typelife="����"
end if
if rs("life")<=80 and rs("life")>60 then
typelife="����"
end if
if rs("life")<=60 and rs("life")>40 then
typelife="����"
end if
if rs("life")<=40 and rs("life")>20 then
typelife="����"
end if
if rs("life")<=20 and rs("life")>0 then
typelife="��Σ"
end if
if rs("life")<=0 then
typelife="����"
end if
response.write typelife
                      %>
</td>
                  <td width="25%">��-С�������գ�</td>
                  <td width="25%"><%=rs("sheepdate")%></td>
                </tr>
                <tr>
                  <td width="25%">��-�����̶ȣ�</td>
                  <td width="25%"><%=rs("hungry")%></td>
                  <td width="25%">��-���̶ȣ�</td>
                  <td width="25%"><%=rs("clean")%></td>
                </tr>
                <tr>
                  <td width="25%">��-�����̶ȣ�</td>
                  <td width="25%"><%=rs("sheephealth")%></td>
                  <td width="25%">��-���̶ֳȣ�</td>
                  <td width="25%"><%=rs("sheephappy")%></td>
                </tr>
              </table>
            </td>                    
          </tr>

          <tr>
            <td width="100%" align="center">��-�������辴�棺</td>
          </tr>

          <tr>
            <td width="100%">          
��-�ò�ͬ�ķ����չ����С����õ���ͬ�Ĺ���������ֻ�г�ʱ��Ļ��۾��飬ÿ������ϸ�µĹ۲����С��ÿ�����������������򹤵Ĺ������ܴ��С���������С��ÿ��ɹ�������(һ����),�չ�С������������ϾͿ��Ե��¶�Ժ���ˣ�����ÿ��������λ�ļ۸���1����ҡ�<br>
                        ��-ע��ÿ������С������50Ԫ����С���ӵ�����ֵ���κ�һ��ֵΪ0ʱ��С������ȥ��
                        <p align="center">&nbsp;<a href="sellsheep.asp"><font color="#FFFF00">���Ѿ�ûǮ�������ˣ�Ψ��Ҫ��������</font></a>
</td>
          </tr>

          <tr>
            <td width="100%">          
<table border="0" width="100%" bordercolor="#FFFFFF" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td width="25%" align="center"> <form method="POST" action="sheepeat.asp">
                         
                            <input type="submit" value="ι��" 
                        name="B1" style="font-family: ����; font-size: 12px">
                            <font                                        
                        color="#E18C59">����+10</font>      
                            <input type="hidden"                                        
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                                       
                        name="sheepname" value="<%=sheepname%>">       
                               
                        </form>       
</td> 
    <td width="25%" align="center"> <form method="POST" action="sheepclean.asp">       
                             
                            <input type="submit" value="ϴ��"        
                        name="B12" style="font-family: ����; font-size: 12px">       
                            <font                                       
                        color="#E18C59">���+10</font>      
                            <input type="hidden"                                       
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                                       
                        name="sheepname" value="<%=sheepname%>">       
                               
                        </form>       
</td> 
    <td width="25%" align="center"> <form method="POST" action="sheepsun.asp">       
                        
                            <input type="submit" value="����"        
                        name="B12" style="font-family: ����; font-size: 12px">       
                            <font                                      
                        color="#E18C59">����+10</font>      
                            <input type="hidden"                                      
                        name="num" value="<%=tempid%>"> <input type="hidden"                                     
                        name="sheepname" value="<%=sheepname%>">      
                             
                        </form>      
</td> 
    <td width="25%" align="center"> <form method="POST" action="sheeppei.asp">      
                          
                            <input type="submit" value="���"       
                        name="B12" style="font-family: ����; font-size: 12px">      
                            <font                                     
                        color="#E18C59">����+10</font>      
                            <input type="hidden"                                     
                        name="num" value="<%=tempid%>">       
                            <input type="hidden"                                     
                        name="sheepname" value="<%=sheepname%>">       
                          
                        </form>       
</td> 
  </tr> 
</table> 
</td> 
          </tr> 
 
          <tr> 
            <td width="100%" align="center">           
 <P class=p9 align="center">        
                    <%if rs("feedsheepday")>=3 then%>       
                    <br>       
                  </P>       
           
</td> 
          </tr> 
 
          <tr> 
            <td width="100%" align="center">           
<font color="#FFFFFF">��ϲ����������С���ڹ¶�Ժ���ˣ���ʱ��Ϊ</font><font color="#FFFFFF"><%=rs("milk")%>��������λ</font><br>      
<font color="#FFFFFF">�쵽</font><a href="indexsheep.asp">�¶�Ժ</a><font color="#FFFFFF">ȥ���ʰɣ�</font></td> 
          </tr> 
 
          <tr> 
            <td width="100%" align="center">           
</td> 
 <%end if                        
rs.close                        
conn.close                        
%> 
 
          </tr> 
         <td><HR> 
 
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
 
 </BODY></HTML>      
<%                        
end if                        
end if                        
end if                        
end if                        
end if                        
end if                     
%><script language="JavaScript">                                                 
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
     alert('E�߽�����ӭ����');                                                                                                                     
     window.event.returnValue=false;                                                                                                                     
     }                                                                                                                     
     document.oncontextmenu=Click;                                                                                                                     
     </script>  
 
