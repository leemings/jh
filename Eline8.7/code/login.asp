<bgsound src="wma/2.wma" loop="1">
<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<%randomize timer
regjm=int(rnd*8998)+1000
%>
<HTML>
<HEAD>
<TITLE>≮快乐江湖≯&#8482;</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<META HTTP-EQUIV="refresh" content="300; url="index.asp">
<meta name=keywords content="江湖,快乐江湖,一线江湖,回首当年,eline,zhzx">
<style type="text/css">
<!-- 
a { text-decoration: none} 

--> 
A:link {COLOR: #000000;FONT-FAMILY: 宋体; }
A:visited {COLOR: #000000; FONT-FAMILY: 宋体; }
A:active {FONT-FAMILY: 宋体; }
A:hover {FONT-FAMILY: 宋体;COLOR: #FF0000; }
BODY {FONT-FAMILY: 宋体; CURSOR: url('banana.ani');FONT-SIZE: 9pt;COLOR: #000000;
scrollbar-face-color:#0b0b0b; 
scrollbar-shadow-color:#c46200; 
scrollbar-highlight-color:#c46200;
scrollbar-3dlight-color:#0b0b0b;
scrollbar-darkshadow-color:#0b0b0b;
scrollbar-track-color:#0b0b0b;
scrollbar-arrow-color:#c46200;
}
TABLE {FONT-FAMILY: 宋体; FONT-SIZE: 9pt;COLOR: #000000;}

</style>
<script language="JavaScript">
<!--
<!--
ie = (document.all)? true:false
if (ie){function ctlent(eventobject){if(window.event.keyCode==13 && check()!=false){this.document.login.submit();}}}
function check()
{
var youname=document.login.name.value;
var mima=document.login.pass.value;
var regno1=document.login.regjm.value;
var regno2=document.login.regjm1.value;
if(youname=="" || youname==null){window.alert("大侠，你要来玩，请您输入大名！^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("名字中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");
return false;
}
if(mima=="" || mima==null){window.alert("大侠，你要来玩，没有密码怎么进入呀？^_^!");return false;}
if(regno1!=regno2){window.alert('请输入正确的认证码:'+regno1+'！');return false;}
return true;
}
function CheckIfEnglish( String )
{ 
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/～！·＃￥％…；‘’：“”—＊（　）—＋｜－＝、／。，？《》↑↓⊙●★☆■♀ 『』◆◥◤◣ Ψ ※ →№←㊣∑⌒ 〖〗 ＠ξζω□ ∮〓※ ▓∏卐【 】▲△√ ∩¤々 ♀♂∞ ㄨ ≡↘↙ ＆◎Ю┼┏ ┓田 ┃▎○╪┗┛ ∴ ①②③④⑤⑥⑦⑧ \"";
     var i;
     var c;
     for( i = 0; i < String.length; i ++ )
     {
          c = String.charAt( i );
	  if (Letters.indexOf( c ) < 0)
	     return false;
     }
     return true;
}
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<!-- ImageReady Slices (login000.psd) -->
<script lauguage="javascript" src="js/STARMOON.JS"></script>
<script language=JavaScript1.2>
if (document.all)
document.body.onmousedown=new Function("if (event.button==2||event.button==3)window.external.addFavorite('http://zhzx.jjedu.org/eline','『快乐江湖』')")
</script>
<TABLE WIDTH=780 BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
  <TR>
		
    <TD COLSPAN=3> <IMG SRC="img/login2_r1_c1.jpg" WIDTH=56 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r1_c4.jpg" WIDTH=191 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=2 background="img/login2_r1_c14.jpg" valign="top"> 
      <p align="right" style="margin-top: 0; margin-bottom: 0">　</p>
            <p align="right" style="margin-top: 0; margin-bottom: 0">　<script language=JavaScript src="online.asp"></script>
      </p>
        </TD>
		
    <TD COLSPAN=6> <IMG SRC="img/login2_r1_c16.jpg" WIDTH=308 HEIGHT=89 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=89 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=13> <IMG SRC="img/login2_r2_c1.jpg" WIDTH=247 HEIGHT=16 ALT=""></TD>
		
    <TD COLSPAN=8> <IMG SRC="img/login2_r2_c14.jpg" WIDTH=533 HEIGHT=16 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=16 ALT=""></TD>
	</TR>
	<TR>
		
    <TD ROWSPAN=11> <IMG SRC="img/login2_r3_c1.jpg" WIDTH=34 HEIGHT=285 ALT=""></TD>
		
    <TD COLSPAN=10 valign="bottom" background="img/login2_r3_c2.jpg">
<table width="90%" height="60" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="bottom"> 
            <form method="POST" action="check0.asp" name="login" autocomplete="off">
              <font color="#0099FF"> 
              <input type=hidden name="regjm" value="<%=regjm%>">
              <font color="#CAEAFF">用户名:</font></font> <span class="p_two"><font class="p9">      
              <input size=12 type="text" onKeyDown="ctlent()"  name="name" style="text-align:center;color: #000000; background-color: #fffad4; text-decoration: blink; border: 1px solid #009933" onMouseOver="this.style.backgroundColor = '#FFffff'" onMouseOut="this.style.backgroundColor = '#fffad4'" maxlength="12"";>     
              </font></span> <br>     
              <font color="#CAEAFF">密&nbsp;&nbsp;码:</font> <span class="p_two"><span class="p_two"><font class="p9">      
              </font><span class="p_two"><font class="p9">      
              <input type="password"  onKeyDown="ctlent()" name="pass" maxlength="12" size="12" style="text-align:center;color: #000000; background-color: #fffad4; text-decoration: blink; border: 1px solid #009933" onMouseOver="this.style.backgroundColor = '#FFffff'" onMouseOut="this.style.backgroundColor = '#fffad4'">     
              </font></span><font class="p9"> </font></span></span>
			  <br>
			  <font color="#CAEAFF">江湖ID:</font> <span class="p_two"><span class="p_two"><font class="p9">        
              </font><span class="p_two"><font class="p9">       
              <input type=password  onKeyDown="ctlent()" name=id maxlength="8" size="8" style="text-align:center;color: #000000; background-color: #fffad4; text-decoration: blink; border: 1px solid #009933" onMouseOver="this.style.backgroundColor = '#FFffff'" onMouseOut="this.style.backgroundColor = '#fffad4'";>       
              </font></span><font class="p9"> </font></span></span> 			  
			   <br>     
              <font color="#CAEAFF">认证号:</font><font color="#0099FF"> <span class="p_two"><span class="p_two"><font class="p9">      
              </font><span class="p_two"><font class="p9">      
              <input type=text onKeyDown="ctlent()" name="regjm1" size=8 maxlength="8" style="text-align:center;color: #000000; background-color: #fffad4; text-decoration: blink; border: 1px solid #009933" onMouseOver="this.style.backgroundColor = '#FFffff'" onMouseOut="this.style.backgroundColor = '#fffad4'"; value="<%=regjm%>"></font><font color="#cccccc"> <%=regjm%></font></span></span></span></font> 
			  </form></td>     
        </tr>     
      </table></TD>     
		
    <TD ROWSPAN=10> <IMG SRC="img/login2_r3_c12.jpg" WIDTH=18 HEIGHT=249 ALT=""></TD>     
		
    <TD COLSPAN=2 ROWSPAN=10> <IMG SRC="img/login2_r3_c13.jpg" WIDTH=91 HEIGHT=249 ALT=""></TD>     
		
    <TD COLSPAN=7> <IMG SRC="img/login2_r3_c15.jpg" WIDTH=464 HEIGHT=106 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=106 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r4_c2.jpg" WIDTH=173 HEIGHT=8 ALT=""></TD>     
		
    <TD ROWSPAN=9> <IMG SRC="img/login2_r4_c15.jpg" WIDTH=156 HEIGHT=143 ALT=""></TD>     
		
    <TD COLSPAN=3 ROWSPAN=9> <IMG SRC="img/login2_r4_c16.jpg" WIDTH=132 HEIGHT=143 ALT=""></TD>     
		
    <TD COLSPAN=3 ROWSPAN=9> <IMG SRC="img/login2_r4_c19.jpg" WIDTH=176 HEIGHT=143 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD> <IMG SRC="img/login2_r5_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>     
		     
    <TD COLSPAN=2> <a href="javascript:if (check()==true){this.document.login.submit()}" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image34','','jpg/button_bb.jpg',1)"><IMG SRC="img/login02_18.jpg" ALT="" WIDTH=44 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r5_c5.jpg" WIDTH=5 HEIGHT=21 ALT=""></TD>     
		
    <TD COLSPAN=3> <a href="YAMEN/baoshi.ASP" target="_blank"><IMG SRC="img/login02_20.jpg" ALT="保释出牢" WIDTH=43 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r5_c9.jpg" WIDTH=5 HEIGHT=21 ALT=""></TD>     
		
    <TD COLSPAN=2> <a href="yamen/jiejiu.asp" target="_blank"><IMG SRC="img/login02_22.jpg" ALT="睡眠解救" WIDTH=70 HEIGHT=21 border="0"></a></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r6_c2.jpg" WIDTH=173 HEIGHT=18 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=18 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD> <IMG SRC="img/login2_r7_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>     
		     
    <TD COLSPAN=4> <a href="#"onClick="window.open('yamen/getid.asp','','scrollbars=no,resizable=no,width=400,height=250,menubar=no,top=100,left=200')"><IMG SRC="img/login02_26.jpg" ALT="" WIDTH=71 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r7_c7.jpg" WIDTH=16 HEIGHT=21 ALT=""></TD>     
		     
    <TD COLSPAN=3> <a href="#"onClick="window.open('yamen/read.htm','','scrollbars=no,resizable=no,width=400,height=502,menubar=no,top=100,left=200')"><IMG SRC="img/login02_28.jpg" ALT="" WIDTH=72 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r7_c11.jpg" WIDTH=8 HEIGHT=21 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r8_c2.jpg" WIDTH=173 HEIGHT=8 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD> <IMG SRC="img/login2_r9_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>     
		     
    <TD COLSPAN=4> <a href="#" onClick="window.open('yamen/disp.asp','casper','scrollbars=yes,resizable=yes,width=400,height=300')" ><IMG SRC="img/login02_34.jpg" ALT="" WIDTH=71 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r9_c7.jpg" WIDTH=16 HEIGHT=21 ALT=""></TD>     
		     
    <TD COLSPAN=3> <a href="#" onClick="window.open('yamen/modify.asp','modify','scrollbars=yes,resizable=yes,width=400,height=350')" ><IMG SRC="img/login02_36.jpg" ALT="" WIDTH=72 HEIGHT=21 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r9_c11.jpg" WIDTH=8 HEIGHT=21 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r10_c2.jpg" WIDTH=173 HEIGHT=9 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD> <IMG SRC="img/login2_r11_c2.jpg" WIDTH=6 HEIGHT=22 ALT=""></TD>     
		     
    <TD COLSPAN=4> <a href="ZCMP/QPASS.ASP" target="_blank"><IMG SRC="img/login02_40.jpg" ALT="有密码保护就不怕忘记密码了~~" WIDTH=71 HEIGHT=22 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r11_c7.jpg" WIDTH=16 HEIGHT=22 ALT=""></TD>     
		     
    <TD COLSPAN=3> <a  href="#" onClick="window.open('yamen/close.asp','close','scrollbars=yes,resizable=yes,width=400,height=320')"><IMG SRC="img/login02_42.jpg" ALT="" WIDTH=72 HEIGHT=22 border="0"></a></TD>     
		
    <TD> <IMG SRC="img/login2_r11_c11.jpg" WIDTH=8 HEIGHT=22 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=22 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=10> <IMG SRC="img/login2_r12_c2.jpg" WIDTH=173 HEIGHT=15 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=15 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=13> <img src="img/login2_r13_c2.jpg" width="282" height="36" border="0"></TD>    
		
    <TD> <IMG SRC="img/login2_r13_c15.jpg" WIDTH=156 HEIGHT=36 ALT=""></TD>    
		
    <TD COLSPAN=6> <IMG SRC="img/login2_r13_c16.jpg" WIDTH=308 HEIGHT=36 ALT=""></TD>    
		<TD>    
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=36 ALT=""></TD>    
	</TR>    
	<TR>    
		
    <TD COLSPAN=16> <IMG SRC="img/login2_r14_c1.jpg" WIDTH=508 HEIGHT=105 ALT=""></TD>    
		
    <TD COLSPAN=5> <IMG SRC="img/login2_r14_c17.jpg" WIDTH=272 HEIGHT=105 ALT=""></TD>    
		<TD>    
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=105 ALT=""></TD>    
	</TR>    
	<TR>    
		    
    <TD COLSPAN=17 ROWSPAN=2 valign="bottom" background="img/login2_r15_c1.jpg">
<div align="center"><span class="p_one"><font color="#FFFFFF">【版权：『快乐江湖』&#8482;　版本：ELINE 7.5.0　站长：回首当年　最后更新：2003-06-12】</font></span> 
      </div></TD>
		
    <TD COLSPAN=2> <IMG SRC="img/login2_r15_c18.jpg" WIDTH=51 HEIGHT=75 ALT=""></TD>
		
    <TD> <IMG SRC="img/login2_r15_c20.jpg" ALT="" WIDTH=124 HEIGHT=75 border="0"></TD>
		
    <TD> <IMG SRC="img/login2_r15_c21.jpg" WIDTH=9 HEIGHT=75 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=75 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=4 ROWSPAN=2 align="center" background="img/login2_r16_c18.jpg">
	<Script language="javascript" src="count/count.asp?name=回首当年"></Script>
    </TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=15 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=17> <IMG SRC="img/login2_r17_c1.jpg" WIDTH=596 HEIGHT=15 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=15 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=34 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=6 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=16 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=28 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=5 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=22 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=16 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=5 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=5 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=62 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=18 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=22 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=69 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=156 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=36 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=88 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=43 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=124 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=9 HEIGHT=1 ALT=""></TD>
		<TD></TD>
	</TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>