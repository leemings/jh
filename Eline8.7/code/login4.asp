<bgsound src="wma/04.wma" loop="1">
<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<%
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=true
shuaxin_time=10
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>��ҳ�������˷�ˢ�»��ƣ��벻Ҫ��<b><font color=ff0000>"&shuaxin_time&"</font></b>��������ˢ�±�ҳ��<BR>���ڴ�ҳ�棬���Ժ򡭡�"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
randomize timer
regjm=int(rnd*8998)+1000
%>
<HTML>
<HEAD>
<TITLE>��E�߽�����վ��&#8482;��ӭ������վ�������� - www.51eline.com</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<META HTTP-EQUIV="refresh" content="200; url="index.asp">
<meta name=keywords content="����,E�߽���,E�߽�����վ,������վ,һ�߽���,һ����,eline,51eline,51eline.com,www.51eline.com"> 
<style type="text/css">
<!-- 
a { text-decoration: none} 

--> 
A:link {COLOR: #000000;FONT-FAMILY: ����; }
A:visited {COLOR: #000000; FONT-FAMILY: ����; }
A:active {FONT-FAMILY: ����; }
A:hover {FONT-FAMILY: ����;COLOR: #FF0000; }
BODY {FONT-FAMILY: ����; CURSOR: url('banana.ani');FONT-SIZE: 9pt;COLOR: #000000;
scrollbar-face-color:#f1d6c1; 
scrollbar-shadow-color:#fcba04; 
scrollbar-highlight-color:#fcba04;
scrollbar-3dlight-color:#f1d6c1;
scrollbar-darkshadow-color:#f1d6c1;
scrollbar-track-color:#f1d6c1;
scrollbar-arrow-color:#fcba04;
}
TABLE {FONT-FAMILY: ����; FONT-SIZE: 9pt;COLOR: #000000;}

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
if(youname=="" || youname==null){window.alert("��������Ҫ���棬�������������^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("�����в����зǷ��ַ���Ӣ�ġ����֣�ֻ��ʹ������Ƭ������^_^!");
return false;
}
if(mima=="" || mima==null){window.alert("��������Ҫ���棬û��������ô����ѽ��^_^!");return false;}
return true;
}
function CheckIfEnglish( String )
{ 
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/�����������������������������������������������������������������ѡ������ ������������ �� �� ������I�ơ� ���� ���ΦƦء� �ӡ��� ���ǅe�� �������� �ɡ衩 ���� �� �ԨK�L ������ੳ ���� ������n���� �� �٢ڢۢܢݢޢߢ� \"";
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

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<!-- ImageReady Slices (login08.psd) -->
<script language=JavaScript1.2>
if (document.all)
document.body.onmousedown=new Function("if (event.button==2||event.button==3)window.external.addFavorite('http://www.51eline.com','��E�߽�����')")
</script>
<TABLE WIDTH=780 BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
  <TR>
		
    <TD COLSPAN=3> <IMG SRC="img/login4_r1_c1.jpg" WIDTH=56 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r1_c4.jpg" WIDTH=191 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=2> <IMG SRC="img/login4_r1_c14.jpg" WIDTH=225 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=6 background="img/login4_r1_c16.jpg"> 
      <p align="center">��<script language=JavaScript src="online.asp"></script></p>
        </TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=89 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=13> <IMG SRC="img/login4_r2_c1.jpg" WIDTH=247 HEIGHT=16 ALT=""></TD>
		
    <TD COLSPAN=8> <IMG SRC="img/login4_r2_c14.jpg" WIDTH=533 HEIGHT=16 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=16 ALT=""></TD>
	</TR>
	<TR>
		
    <TD ROWSPAN=11> <IMG SRC="img/login4_r3_c1.jpg" WIDTH=34 HEIGHT=285 ALT=""></TD>
		
    <TD COLSPAN=10 background="img/login4_r3_c2.jpg">
<table width="90%" height="80" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> 
            <form method="POST" action="check0.asp" name="login" autocomplete="off">
              <font color="#0099FF"> 
              <input type=hidden name="regjm" value="<%=regjm%>">
              <font color="#CAEAFF">�û���:</font></font> <span class="p_two"><font class="p9"> 
              <input size=12 type="text" onKeyDown="ctlent()"  name="name" style="text-align:center;color: #000000; background-color: #ffffff; text-decoration: blink; border: 1px solid #000000" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#ffffff'" maxlength="12"";>
              </font></span> <br>       
              <font color="#CAEAFF">��&nbsp;&nbsp;��:</font> <span class="p_two"><span class="p_two"><font class="p9">        
              </font><span class="p_two"><font class="p9">       
              <input type="password"  onKeyDown="ctlent()" name="pass" maxlength="12" size="12" style="text-align:center;color: #000000; background-color: #ffffff; text-decoration: blink; border: 1px solid #000000" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#ffffff'" solid">       
              </font></span><font class="p9"> </font></span></span>
              <br>
              <font color="#CAEAFF">��&nbsp;&nbsp;��:</font> <select name=txwz style=background-color:#ffffff;font-size:9pt;color:000000>
                    <option value="0" selected>������</option>
                    <option value="1">�� ֮ ��</option>
                    <option value="2">����֮��</option>
                    <option value="3">������԰</option>
                    <option value="4">ʥ���޵�</option>
                    <option value="5">һ��ʹ��</option>
                    <option value="6">���˴�˵</option>
                    <option value="7">��ʿ����</option>
                    <option value="8">ͯ������</option>
                    <option value="9">�ɰ�����</option>
                    <option value="10">ħ��ɭ��</option>
                    <option value="11">��������</option>
                  </select> 
            </form></td>       
        </tr>       
      </table> </TD>       
		
    <TD ROWSPAN=10> <IMG SRC="img/login4_r3_c12.jpg" WIDTH=18 HEIGHT=249 ALT=""></TD>       
		
    <TD COLSPAN=2 ROWSPAN=10> <IMG SRC="img/login4_r3_c13.jpg" WIDTH=91 HEIGHT=249 ALT=""></TD>       
		
    <TD COLSPAN=7> <IMG SRC="img/login4_r3_c15.jpg" WIDTH=464 HEIGHT=106 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=106 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r4_c2.jpg" WIDTH=173 HEIGHT=8 ALT=""></TD>       
		
    <TD ROWSPAN=9> <IMG SRC="img/login4_r4_c15.jpg" WIDTH=156 HEIGHT=143 ALT=""></TD>       
		
    <TD COLSPAN=3 ROWSPAN=9> <IMG SRC="img/login4_r4_c16.jpg" WIDTH=132 HEIGHT=143 ALT=""></TD>       
		
    <TD COLSPAN=3 ROWSPAN=9> <IMG SRC="img/login4_r4_c19.jpg" WIDTH=176 HEIGHT=143 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD> <IMG SRC="img/login4_r5_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>       
		       
    <TD COLSPAN=2> <a href="javascript:if (check()==true){this.document.login.submit()}" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image34','','jpg/button_bb.jpg',1)"><IMG SRC="img/login02_18.jpg" ALT="" WIDTH=44 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r5_c5.jpg" WIDTH=5 HEIGHT=21 ALT=""></TD>       
		
    <TD COLSPAN=3> <a href="YAMEN/baoshi.ASP" target="_blank"><IMG SRC="img/login02_20.jpg" ALT="���ͳ���" WIDTH=43 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r5_c9.jpg" WIDTH=5 HEIGHT=21 ALT=""></TD>       
		
    <TD COLSPAN=2> <a href="yamen/jiejiu.asp" target="_blank"><IMG SRC="img/login02_22.jpg" ALT="˯�߽��" WIDTH=70 HEIGHT=21 border="0"></a></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r6_c2.jpg" WIDTH=173 HEIGHT=18 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=18 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD> <IMG SRC="img/login4_r7_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>       
		       
    <TD COLSPAN=4> <a href="#"onClick="window.open('yamen/getid.asp','','scrollbars=no,resizable=no,width=400,height=250,menubar=no,top=100,left=200')"><IMG SRC="img/login02_26.jpg" ALT="" WIDTH=71 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r7_c7.jpg" WIDTH=16 HEIGHT=21 ALT=""></TD>       
		       
    <TD COLSPAN=3> <a href="#"onClick="window.open('yamen/read.htm','','scrollbars=no,resizable=no,width=400,height=502,menubar=no,top=100,left=200')"><IMG SRC="img/login02_28.jpg" ALT="" WIDTH=72 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r7_c11.jpg" WIDTH=8 HEIGHT=21 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r8_c2.jpg" WIDTH=173 HEIGHT=8 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD> <IMG SRC="img/login4_r9_c2.jpg" WIDTH=6 HEIGHT=21 ALT=""></TD>       
		       
    <TD COLSPAN=4> <a href="#" onClick="window.open('yamen/disp.asp','casper','scrollbars=yes,resizable=yes,width=400,height=300')" ><IMG SRC="img/login02_34.jpg" ALT="" WIDTH=71 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r9_c7.jpg" WIDTH=16 HEIGHT=21 ALT=""></TD>       
		       
    <TD COLSPAN=3> <a href="#" onClick="window.open('yamen/modify.asp','modify','scrollbars=yes,resizable=yes,width=400,height=350')" ><IMG SRC="img/login02_36.jpg" ALT="" WIDTH=72 HEIGHT=21 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r9_c11.jpg" WIDTH=8 HEIGHT=21 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r10_c2.jpg" WIDTH=173 HEIGHT=9 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD> <IMG SRC="img/login4_r11_c2.jpg" WIDTH=6 HEIGHT=22 ALT=""></TD>       
		       
    <TD COLSPAN=4> <a href="ZCMP/QPASS.ASP" target="_blank"><IMG SRC="img/login02_40.jpg" ALT="�����뱣���Ͳ�������������~~" WIDTH=71 HEIGHT=22 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r11_c7.jpg" WIDTH=16 HEIGHT=22 ALT=""></TD>       
		       
    <TD COLSPAN=3> <a  href="#" onClick="window.open('yamen/close.asp','close','scrollbars=yes,resizable=yes,width=400,height=320')"><IMG SRC="img/login02_42.jpg" ALT="" WIDTH=72 HEIGHT=22 border="0"></a></TD>       
		
    <TD> <IMG SRC="img/login4_r11_c11.jpg" WIDTH=8 HEIGHT=22 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=22 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=10> <IMG SRC="img/login4_r12_c2.jpg" WIDTH=173 HEIGHT=15 ALT=""></TD>       
		<TD>       
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=15 ALT=""></TD>       
	</TR>       
	<TR>       
		
    <TD COLSPAN=13> <img src="img/login4_r13_c2.jpg" width="282" height="36" border="0"></TD>     
		
    <TD> <IMG SRC="img/login4_r13_c15.jpg" WIDTH=156 HEIGHT=36 ALT=""></TD>     
		
    <TD COLSPAN=6> <IMG SRC="img/login4_r13_c16.jpg" WIDTH=308 HEIGHT=36 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=36 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=16> <IMG SRC="img/login4_r14_c1.jpg" WIDTH=508 HEIGHT=105 ALT=""></TD>     
		
    <TD COLSPAN=5> <IMG SRC="img/login4_r14_c17.jpg" WIDTH=272 HEIGHT=105 ALT=""></TD>     
		<TD>     
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=105 ALT=""></TD>     
	</TR>     
	<TR>     
		
    <TD COLSPAN=17 ROWSPAN=2 background="img/login4_r15_c1.jpg" valign="bottom"> 
      ��<font color="#FFFFFF">����Ȩ��<a href="http://www.51eline.com"><font color=#ffffff>��һ�����硻&#8482;</font></a> ��Ȩ����<%=Application("sjjh_chatroomname")%>�� �汾��ELINE V8.7���ް桡վ����<%=Application("sjjh_admin")%>��</font></TD>    
		
    <TD COLSPAN=2> <IMG SRC="img/login4_r15_c18.jpg" WIDTH=51 HEIGHT=75 ALT=""></TD>    
		
    <TD> <img src="img/login4_r15_c20.jpg" width="124" height="75" border="0"></TD>    
		
    <TD> <IMG SRC="img/login4_r15_c21.jpg" WIDTH=9 HEIGHT=75 ALT=""></TD>    
		<TD>    
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=75 ALT=""></TD>    
	</TR>    
	<TR>    
		
    <TD COLSPAN=4 ROWSPAN=2 align="left" background="img/login4_r16_c18.jpg"> 
      
    </TD>    
		<TD>    
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=15 ALT=""></TD>    
	</TR>    
	<TR>    
		
    <TD COLSPAN=17> <IMG SRC="img/login4_r17_c1.jpg" WIDTH=596 HEIGHT=15 ALT=""></TD>    
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