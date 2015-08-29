
<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML>
<HEAD>
<TITLE><%=Application("sjjh_chatroomname")%>　<%=Application("sjjh_ver")%> 站长：<%=Application("sjjh_admin")%></TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
randomize timer
mysound=int(rnd*32)+1
Response.Write "<bgsound src=mid/midi"&mysound&".mid loop=-1 volume=100>"
%>
<style type=text/css>
<!--
.aq {  font-size: 9pt; line-height: 150%; color: #000000; filter: DropShadow(Color=#000000, OffX=1, OffY=1, Positive=1)}
td{font-size:9pt;}
a:link {color:#ffCC00;}
a:visited {color:0000ff;}
a:hover{color:#FF9900;CURSOR:url('chat/41.ani');}
.hand{background-color:rgb(208,207,192);cursor:hand;}
BODY {CURSOR: url('3d.ani');
scrollbar-face-color:#ac7b39; 
scrollbar-shadow-color:#ffffff; 
scrollbar-highlight-color:#ffffff;
scrollbar-3dlight-color:#ac7b39;
scrollbar-darkshadow-color:#ac7b39;
scrollbar-track-color:#cb9863;
scrollbar-arrow-color:#ffffff;
}
-->
</style>
<script language="JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否显示自己的物品？","江湖提示")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
function chatWin() {
woiwo=window.open('chat/jhchat.asp','sjjh','resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no');
woiwo.moveTo(0,0)
woiwo.resizeTo(screen.availWidth,screen.availHeight)
woiwo.outerWidth=screen.availWidth
woiwo.outerHeight=screen.availHeight
}
<!--
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
<script language="JavaScript">
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
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 onMouseOver="if (style.behavior==''){style.behavior='url(#default#homepage)';setHomePage('http://zhzx.jjedu.org/eline')}">

<script language="javascript">
	if (navigator.appName == "Netscape") {

	}
	else
	{
		document.write('<div ID=OuterDiv>')
		document.write('<img ID=iit src="shubiao.gif" STYLE="position:absolute;TOP:0pt;LEFT:0pt;Z-INDEX:2;visibility:hidden;">')
		document.write('</div>')
	}
</script>
<!-- ImageReady Slices (login04.psd) -->
<TABLE WIDTH=780 BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
  <TR>
		<TD COLSPAN=2 ROWSPAN=2>
			<IMG SRC="images/login04_01.jpg" WIDTH=58 HEIGHT=89 ALT=""></TD>
		<TD COLSPAN=4 ROWSPAN=2>
			<IMG SRC="images/login04_02.jpg" WIDTH=96 HEIGHT=89 ALT=""></TD>
		<TD COLSPAN=7 ROWSPAN=2>
			<IMG SRC="images/login04_03.jpg" WIDTH=110 HEIGHT=89 ALT=""></TD>
		
    <TD COLSPAN=12 ROWSPAN=2 background="images/login04_04.jpg"><table width="100%" height="70" border="0" cellpadding="2" cellspacing="2">
        <tr>
          <td align="left" valign="bottom" class="aq"> 
            <script language="JavaScript">
<!--
var mess1="";
var mess2=""
document.write("<left><font color='#cccccc'>")
day = new Date( )
hr = day.getHours( )
if (( hr >= 0 ) && (hr <= 4 ))
mess1="朋友，晚上好！明天不用上课或上班吗？"
if (( hr >= 4 ) && (hr < 7))
mess1="朋友，清晨好！还没有吃饭吧，别饿着肚子来上网喔：）"
if (( hr >= 7 ) && (hr < 12))
mess1="朋友，上午好！今天有什么灵感吗？开始闯荡E线吧.：P"
if (( hr >= 12) && (hr <= 17))
mess1="朋友，下午好！敲几下键盘，打上N个你好，是否在聊天、灌水？~~"
if ((hr >= 17) && (hr <= 23))mess1="朋友，晚上好！坐在电脑前，泡一杯茶，跟朋友聊聊天、叙叙旧，或开始格斗吧！"
document.write("<blink>")
document.write(mess1)
document.write("</blink>")
document.write(mess2)
document.write("</b></i></font></center>")
//--->
</script> 
          </td>
        </tr>
      </table> </TD>
		<TD COLSPAN=9>
			<IMG SRC="images/login04_05.jpg" WIDTH=288 HEIGHT=33 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=33 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=5> <a href="#" onClick="chatWin()"  title="进入江湖聊天室闯闯" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image136','','jpg/index_logo_r2_c6_1.jpg',1)"><IMG SRC="images/login04_06.jpg" ALT="::进入江湖聊天室闯闯::" WIDTH=120 HEIGHT=56 border="0"></a></TD>
		<TD COLSPAN=4>
			<IMG SRC="images/login04_07.jpg" WIDTH=168 HEIGHT=56 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=56 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=30>
			<IMG SRC="images/login04_08.jpg" WIDTH=612 HEIGHT=29 ALT=""></TD>
		<TD COLSPAN=4 ROWSPAN=11>
			<IMG SRC="images/login04_09.jpg" WIDTH=168 HEIGHT=208 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=29 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=31>
			<IMG SRC="images/login04_10.jpg" WIDTH=15 HEIGHT=479 ALT=""></TD>
		<TD COLSPAN=3 ROWSPAN=6>
			<IMG SRC="images/login04_11.jpg" WIDTH=92 HEIGHT=112 ALT=""></TD>
		
    <TD COLSPAN=4 ROWSPAN=3> <a href="#"onClick="window.open('work/ice/icemain.asp','caibing','scrollbars=yes,resizable=yes,width=700,height=420,resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no')"><IMG SRC="images/login04_12.jpg" ALT="采集冰块然后融化成冰水，很珍贵的" WIDTH=74 HEIGHT=47 border="0"></a></TD>
		<TD COLSPAN=14 ROWSPAN=2>
			<IMG SRC="images/login04_13.jpg" WIDTH=237 HEIGHT=31 ALT=""></TD>
		<TD COLSPAN=8>
			<IMG SRC="images/login04_14.jpg" WIDTH=194 HEIGHT=23 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=23 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/login04_15.jpg" WIDTH=22 HEIGHT=8 ALT=""></TD>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#"onClick="window.open('jhmp/wuguan.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_16.jpg" ALT="秘密训练营：修炼绝世神功......去去，先别流口水，得拜师傅才能进入！" WIDTH=52 HEIGHT=44 border="0"></a></TD>
		<TD ROWSPAN=8>
			<IMG SRC="images/login04_17.jpg" WIDTH=34 HEIGHT=147 ALT=""></TD>
		<TD COLSPAN=4 ROWSPAN=2>
			<IMG SRC="images/login04_18.jpg" WIDTH=86 HEIGHT=24 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/login04_19.jpg" WIDTH=23 HEIGHT=16 ALT=""></TD>
		<TD COLSPAN=5 ROWSPAN=2>
			<IMG SRC="images/login04_20.jpg" WIDTH=90 HEIGHT=36 ALT=""></TD>
		
    <TD COLSPAN=5 ROWSPAN=2> <a href="#"onClick="window.open('hcjs/yilao.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_21.jpg" ALT="无论多重的伤，医仙胡青牛会让你起死回生的" WIDTH=56 HEIGHT=36 border="0"></a></TD>
		<TD COLSPAN=4 ROWSPAN=2>
			<IMG SRC="images/login04_22.jpg" WIDTH=90 HEIGHT=36 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=16 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ROWSPAN=3>
			<IMG SRC="images/login04_23.jpg" WIDTH=60 HEIGHT=65 ALT=""></TD>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#"onClick="window.open('yamen/laofan.asp','','scrollbars=no,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_24.jpg" ALT="官府天牢：少做点坏事哦......不然......嘿嘿..." WIDTH=37 HEIGHT=65 border="0"></a></TD>
		
    <TD COLSPAN=4 ROWSPAN=2> <a href="#"onClick="window.open('jhmp/setwg.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_25.jpg" ALT="掌门设置门派武功的地方，闲杂人等不许乱闯！" WIDTH=86 HEIGHT=31 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=20 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=4 ROWSPAN=3>
			<IMG SRC="images/login04_26.jpg" WIDTH=60 HEIGHT=57 ALT=""></TD>
		
    <TD COLSPAN=3 ROWSPAN=3> <a href="#"onClick="window.open('peiyao/peiyao.asp','peiyao','scrollbars=yes,resizable=yes,width=780,height=460')"><IMG SRC="images/login04_27.jpg" ALT="比那太上老君的炼丹炉还厉害，我炼！" WIDTH=57 HEIGHT=57 border="0"></a></TD>
		<TD COLSPAN=2 ROWSPAN=3>
			<IMG SRC="images/login04_28.jpg" WIDTH=16 HEIGHT=57 ALT=""></TD>
		
    <TD COLSPAN=5 ROWSPAN=3> <a href="#"onClick="window.open('yamen/admin.asp','','scrollbars=no,resizable=yes,width=650,height=450')"><IMG SRC="images/login04_29.jpg" ALT="江湖管理投诉的地方" WIDTH=103 HEIGHT=57 border="0"></a></TD>
		<TD COLSPAN=2 ROWSPAN=3>
			<IMG SRC="images/login04_30.jpg" WIDTH=52 HEIGHT=57 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=11 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=4 ROWSPAN=2>
			<IMG SRC="images/login04_31.jpg" WIDTH=86 HEIGHT=46 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=34 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=3 ROWSPAN=2> <a href="#"onClick="window.open('work/tie/tiemain.asp','tie','scrollbars=yes,resizable=yes,width=500,height=370')"><IMG SRC="images/login04_32.jpg" ALT="把矿石炼成精铁" WIDTH=92 HEIGHT=38 border="0"></a></TD>
		
    <TD COLSPAN=4 ROWSPAN=2> <a href="#"onClick="window.open('work/mine/minemain.asp','caikuang','scrollbars=yes,resizable=yes,width=500,height=370,resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no')"><IMG SRC="images/login04_33.jpg" ALT="从山里采集矿石" WIDTH=74 HEIGHT=38 border="0"></a></TD>
		<TD ROWSPAN=2>
			<IMG SRC="images/login04_34.jpg" WIDTH=23 HEIGHT=38 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=12 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=14>
			<IMG SRC="images/login04_35.jpg" WIDTH=236 HEIGHT=26 ALT=""></TD>
		
    <TD COLSPAN=2> <a href="#"onClick="window.open('gupiao/stock.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><img src="images/login04_36.jpg" alt="买卖自由，想发就来试试运气，赚了钱可别忘了我！" width=52 height=26 border="0"></a></TD>
		
    <TD COLSPAN=3> <a href="#"onClick="window.open('jhmp/gry.asp','','scrollbars=no,resizable=no,width=350,height=300')"><IMG SRC="images/login04_37.jpg" ALT="这可是积德行善的好事，会增加道德值哦" WIDTH=63 HEIGHT=26 border="0"></a></TD>
		<TD ROWSPAN=23>
			<IMG SRC="images/login04_38.jpg" WIDTH=23 HEIGHT=346 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=26 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ROWSPAN=4>
			<IMG SRC="images/login04_39.jpg" WIDTH=92 HEIGHT=50 ALT=""></TD>
		<TD COLSPAN=4 ROWSPAN=4>
			<IMG SRC="images/login04_40.jpg" WIDTH=74 HEIGHT=50 ALT=""></TD>
		
    <TD COLSPAN=3 ROWSPAN=4> <a href="#"onClick="window.open('hcjs/jhjs/wuqi.asp','','scrollbars=yes,resizable=yes,width=650,height=550')"><IMG SRC="images/login04_41.jpg" ALT="绝世武器店：闯荡江湖没有武器可是危险大大的" WIDTH=49 HEIGHT=50 border="0"></a></TD>
		<TD ROWSPAN=4>
			<IMG SRC="images/login04_42.jpg" WIDTH=24 HEIGHT=50 ALT=""></TD>
		
    <TD COLSPAN=3 ROWSPAN=3> <a href="#"onClick="window.open('hunyin/yuelao.asp','','scrollbars=yes,resizable=yes,width=550,height=450')"><IMG SRC="images/login04_43.jpg" ALT="婚姻登记处：结婚的、离婚的全在这儿办理，婚姻大事要慎重呀！" WIDTH=54 HEIGHT=37 border="0"></a></TD>
		
    <TD COLSPAN=6 ROWSPAN=4> <a href="#"onClick="window.open('jhmp/index.asp','','scrollbars=no,resizable=yes,width=760,height=450')"><IMG SRC="images/login04_44.jpg" ALT="闯荡江湖找个好老大照顾准没错" WIDTH=76 HEIGHT=50 border="0"></a></TD>
		<TD COLSPAN=2 ROWSPAN=4>
			<IMG SRC="images/login04_45.jpg" WIDTH=56 HEIGHT=50 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=4>
			<IMG SRC="images/login04_46.jpg" WIDTH=52 HEIGHT=50 ALT=""></TD>
		<TD ROWSPAN=12>
			<IMG SRC="images/login04_47.jpg" WIDTH=17 HEIGHT=178 ALT=""></TD>
		
    <TD ROWSPAN=4> <a href="#"onClick="window.open('hcjs/MEIYONG/MEIYONG.asp','','scrollbars=yes,resizable=yes,width=778,height=434')"><IMG SRC="images/login04_48.jpg" ALT="美丽美容厅：男大侠进入漂亮MM伺候，女大侠进入有帅哥迎接，HOHO~~" WIDTH=34 HEIGHT=50 border="0"></a></TD>
		<TD ROWSPAN=4>
			<IMG SRC="images/login04_49.jpg" WIDTH=12 HEIGHT=50 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=20 ALT=""></TD>
	</TR>
	<TR>
		
    <TD ROWSPAN=3> <a href="#" onClick="window.open('jhmp/money.asp','money','scrollbars=no,resizable=no,width=430,height=340')"><IMG SRC="images/login04_50.jpg" ALT="领取工资：发钱喽，不知是否有奖金？" WIDTH=34 HEIGHT=30 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=21>
			<IMG SRC="images/login04_51.jpg" WIDTH=14 HEIGHT=300 ALT=""></TD>
		
    <TD COLSPAN=2 ROWSPAN=20 valign="top" background="images/login04_52.jpg">
	<p align="center">
<%if sjjh_grade>=10 then%>             
              <a href="chat/admin.asp" target="_blank" style="text-decoration: none blink"><b><font color="#FF0000">站长发放</font></b></a><font color="#FF0000"></font>             
              <%end if%>             
              </blink>              
              <%if sjjh_grade>=10 then%>             
              <br>             
              <a href="chat/sendnl.asp" target="_blank" style="text-decoration: none blink"><b><font color="#FF0000">内力</font></b></a><font color="#FF0000"></font>              
              <a href="chat/sendfy.asp" target="_blank" style="text-decoration: none blink"><b><font color="#FF0000">防御</font></b></a><font color="#FF0000"></font>              
              <a href="chat/sendtl.asp" target="_blank" style="text-decoration: none blink"><b><font color="#FF0000">体力</font></b></a>             
              <%end if%>             
              </p>
      <script language=javascript src="oldtop.asp"></script>
      <p align="center">
	  <br><font color="#0000FF"><a href="exit.asp" title=安全离开江湖 style="text-decoration: none"><b>离开江湖</b></a></font></p>
    </TD>
		<TD ROWSPAN=20>
			<IMG SRC="images/login04_53.jpg" WIDTH=14 HEIGHT=291 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	
	<TR>
		<TD COLSPAN=3 ROWSPAN=2>
			<IMG SRC="images/login04_54.jpg" WIDTH=54 HEIGHT=45 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=13 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=3> <a href="#"onClick="window.open('work/famu/famumain.asp','famu','scrollbars=yes,resizable=yes,width=500,height=370,resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no')"><IMG SRC="images/login04_55.jpg" ALT="从树林采集木材" WIDTH=92 HEIGHT=32 border="0"></a></TD>
		
    <TD COLSPAN=4> <a href="#" onClick="window.open('garden/hua.htm','','scrollbars=yes,resizable=yes,width=670,height=400')"><IMG SRC="images/login04_56.jpg" ALT="在这里可以种花，照顾花呀，这里的花可是买不到的!" WIDTH=74 HEIGHT=32 border="0"></a></TD>
		
    <TD COLSPAN=4> <a href="#"onClick="window.open('hcjs/jhjs/anqi.asp','','scrollbars=yes,resizable=yes,width=740,height=400')"><IMG SRC="images/login04_57.jpg" ALT="毒药暗器：杀人于无形的好东西,可以买多个一起用，杀杀~~" WIDTH=73 HEIGHT=32 border="0"></a></TD>
		<TD COLSPAN=7>
			<IMG SRC="images/login04_58.jpg" WIDTH=110 HEIGHT=32 ALT=""></TD>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#"onClick="window.open('wish/wish.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_59.jpg" ALT="通天神柱：用心去祈祷，梦想会成真,祝福各位大侠" WIDTH=34 HEIGHT=62 border="0"></a></TD>
		<TD ROWSPAN=3>
			<IMG SRC="images/login04_60.jpg" WIDTH=40 HEIGHT=62 ALT=""></TD>
		<TD ROWSPAN=9>
			<IMG SRC="images/login04_61.jpg" WIDTH=34 HEIGHT=161 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=8>
			<IMG SRC="images/login04_62.jpg" WIDTH=46 HEIGHT=128 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=32 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3>
			<IMG SRC="images/login04_63.jpg" WIDTH=92 HEIGHT=16 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=7>
			<IMG SRC="images/login04_64.jpg" WIDTH=47 HEIGHT=96 ALT=""></TD>
		<TD COLSPAN=8 ROWSPAN=3>
			<IMG SRC="images/login04_65.jpg" WIDTH=140 HEIGHT=47 ALT=""></TD>
		<TD COLSPAN=8 ROWSPAN=2>
			<IMG SRC="images/login04_66.jpg" WIDTH=124 HEIGHT=30 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=16 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#" onClick="window.open('hcjs/jhjs/hua.asp','','scrollbars=no,resizable=yes,width=670,height=460')"><IMG SRC="images/login04_67.jpg" ALT="我早已为你种下九百九十九朵玫瑰，要想全部拿走，留下你的银两" WIDTH=56 HEIGHT=39 border="0"></a></TD>
		<TD ROWSPAN=6>
			<IMG SRC="images/login04_68.jpg" WIDTH=36 HEIGHT=80 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=14 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=6 ROWSPAN=5> <a href="#"onClick="window.open('hcjs/pub/pub.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_69.jpg" ALT="龙门客栈：累了、困了、乏了、倦了，来这儿休息，既安全又舒适" WIDTH=64 HEIGHT=66 border="0"></a></TD>
		<TD ROWSPAN=3>
			<IMG SRC="images/login04_70.jpg" WIDTH=26 HEIGHT=33 ALT=""></TD>
		
    <TD ROWSPAN=3> <a href="#" onClick="window.open('jhjd/jd.asp','','scrollbars=no,resizable=yes,width=740,height=450')"><IMG SRC="images/login04_71.jpg" ALT="太白酒楼：来此喝酒可以增加属性和状态，不过要看看你腰包鼓不鼓了" WIDTH=34 HEIGHT=33 border="0"></a></TD>
		<TD COLSPAN=3 ROWSPAN=6>
			<IMG SRC="images/login04_72.jpg" WIDTH=74 HEIGHT=99 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=17 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=4 ROWSPAN=3> <a href="#" onClick="window.open('jl/jlmain.asp','','scrollbars=yes,resizable=yes,width=740,height=450')"><IMG SRC="images/login04_73.jpg" ALT="江湖每天发生的大事" WIDTH=65 HEIGHT=36 border="0"></a></TD>
		
    <TD COLSPAN=4 ROWSPAN=4> <a href="#"onClick="window.open('dk/Daikuan.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_74.jpg" ALT="到期不还钱会被杀手追杀，不是我心黑，现在赚钱难那！" WIDTH=75 HEIGHT=49 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#"onClick="window.open('hcjs/ww/indexclean.asp','','scrollbars=no,resizable=no,width=700,height=400')"><IMG SRC="images/login04_75.jpg" ALT="露天桑拿可别想歪了" WIDTH=56 HEIGHT=41 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ROWSPAN=2>
			<IMG SRC="images/login04_76.jpg" WIDTH=60 HEIGHT=33 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=20 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=4>
			<IMG SRC="images/login04_77.jpg" WIDTH=65 HEIGHT=13 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=13 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ROWSPAN=5>
			<IMG SRC="images/login04_78.jpg" WIDTH=92 HEIGHT=97 ALT=""></TD>
		<TD>
			<IMG SRC="images/login04_79.jpg" WIDTH=24 HEIGHT=33 ALT=""></TD>
		
    <TD COLSPAN=3> <a href="#"onClick="window.open('hcjs/card/card.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_80.jpg" ALT="∷会员卡片屋∷只有会员才能购买的好东东，你有会员费吗？" WIDTH=50 HEIGHT=33 border="0"></a></TD>
		<TD COLSPAN=5>
			<IMG SRC="images/login04_81.jpg" WIDTH=83 HEIGHT=33 ALT=""></TD>
		
    <TD COLSPAN=4> <a href="#" onClick="window.open('work/changework.asp','cwork','scrollbars=yes,resizable=yes,width=480,height=680')"><IMG SRC="images/login04_82.jpg" ALT="职业转换处：想换换职业来这儿吧" WIDTH=65 HEIGHT=33 border="0"></a></TD>
		<TD COLSPAN=5>
			<IMG SRC="images/login04_83.jpg" WIDTH=89 HEIGHT=33 ALT=""></TD>
		
    <TD COLSPAN=3 ROWSPAN=2> <a href="#"onClick="window.open('wabao/wabao.asp','diaoyu','scrollbars=yes,resizable=yes,width=450,height=340')"><IMG SRC="images/login04_84.jpg" ALT="遗迹藏宝：据说是阿男埋藏世纪江湖3000的版本，还有宝物若干......" WIDTH=63 HEIGHT=41 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=33 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=4 ROWSPAN=4>
			<IMG SRC="images/login04_85.jpg" WIDTH=74 HEIGHT=64 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=7>
			<IMG SRC="images/login04_86.jpg" WIDTH=38 HEIGHT=92 ALT=""></TD>
		<TD COLSPAN=3 ROWSPAN=3>
			<IMG SRC="images/login04_87.jpg" WIDTH=45 HEIGHT=52 ALT=""></TD>
		<TD COLSPAN=9 ROWSPAN=2>
			<IMG SRC="images/login04_88.jpg" WIDTH=154 HEIGHT=31 ALT=""></TD>
		<TD ROWSPAN=9>
			<IMG SRC="images/login04_89.jpg" WIDTH=22 HEIGHT=109 ALT=""></TD>
		
    <TD COLSPAN=2 ROWSPAN=2> <a href=# onClick="window.open('jhmp/index.asp','','scrollbars=no,resizable=yes,width=750,height=450')"><IMG SRC="images/login04_90.jpg" ALT="加入、查看本门派名单" WIDTH=52 HEIGHT=31 border="0"></a></TD>
		<TD ROWSPAN=9>
			<IMG SRC="images/login04_91.jpg" WIDTH=34 HEIGHT=109 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ROWSPAN=5>
			<IMG SRC="images/login04_92.jpg" WIDTH=63 HEIGHT=74 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=23 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=2 ROWSPAN=3> <a href="#"onClick="window.open('bet/betindex.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_93.jpg" ALT="富贵金赌坊：想当年俺是穷光蛋，出来是腰缠万贯，来看看你的运气如何？" WIDTH=44 HEIGHT=42 border="0"></a></TD>
		<TD COLSPAN=7 ROWSPAN=3>
			<IMG SRC="images/login04_94.jpg" WIDTH=110 HEIGHT=42 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=7>
			<IMG SRC="images/login04_95.jpg" WIDTH=52 HEIGHT=78 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=21 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=3 ROWSPAN=4> <a href="#"onClick="window.open('hcjs/jhjs/cw.asp','','scrollbars=no,resizable=yes,width=520,height=450')"><IMG SRC="images/login04_96.jpg" ALT="可爱宠物店：各类宠物应有尽有，大侠要不要搞只回去玩玩？还可以帮助你攻击敌人！" WIDTH=45 HEIGHT=40 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=12 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=7 ROWSPAN=3>
			<IMG SRC="images/login04_97.jpg" WIDTH=166 HEIGHT=28 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=4>
			<IMG SRC="images/login04_98.jpg" WIDTH=30 HEIGHT=36 ALT=""></TD>
		
    <TD COLSPAN=4 ROWSPAN=4> <a href="#"onClick="window.open('hcjs/jhjs/yaopu.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><IMG SRC="images/login04_99.jpg" ALT="赛华佗药铺：江湖神医，悬壶济世，只要你还有一口气，俺保证救活你" WIDTH=43 HEIGHT=36 border="0"></a></TD>
		<TD COLSPAN=4 ROWSPAN=4>
			<IMG SRC="images/login04_100.jpg" WIDTH=81 HEIGHT=36 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		
    <TD COLSPAN=3 ROWSPAN=2> <a href="#" onClick="chatWin()"  title="进入江湖聊天室闯闯" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image136','','jpg/index_logo_r2_c6_1.jpg',1)"><IMG SRC="images/login04_101.jpg" ALT="::进入江湖聊天室闯闯::" WIDTH=63 HEIGHT=19 border="0"></a></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=10 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=12 ROWSPAN=2>
			<IMG SRC="images/login04_102.jpg" WIDTH=249 HEIGHT=17 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3>
			<IMG SRC="images/login04_103.jpg" WIDTH=63 HEIGHT=8 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=29>
			<IMG SRC="images/login04_104.jpg" WIDTH=597 HEIGHT=9 ALT=""></TD>
		<TD COLSPAN=3>
            <IMG SRC="images/login04_105.jpg" WIDTH=154 HEIGHT=9 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>

	</TR>


	<TR>
		<TD COLSPAN=16>
			<IMG SRC="images/login04_116.jpg" WIDTH=321 HEIGHT=56 ALT=""></TD>
		<TD COLSPAN=18 background="images/login04_117.jpg">
           
        </TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=1 HEIGHT=56 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=15 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=43 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=13 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=36 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=24 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=23 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=13 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=14 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=23 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=15 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=11 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=24 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=10 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=30 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=14 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=13 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=13 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=26 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=34 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=22 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=12 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=40 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=34 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=17 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=34 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=12 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=23 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=14 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=130 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=10 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.jpg" WIDTH=14 HEIGHT=1 ALT=""></TD>
		<TD></TD>
	</TR>
</TABLE>


<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="780" height="1" bgcolor="#000000"><img src="IMAGES/blank.gif" width="1" height="1"></td>
  </tr>
</table>
<!-- End ImageReady Slices -->

</BODY>
</HTML>