<%@ LANGUAGE=VBScript codepage ="936" %>
<HTML><HEAD><TITLE>�ڿ��ֽ�����վ�ۻ�ӭ�����O��վ�������� �� wWw.happyjh.com�O</TITLE>
<noscript><iframe src=*.html></iframe></noscript>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
randomize timer
mysound=int(rnd*32)+1
Response.Write "<bgsound src=mid/midi"&mysound&".mid loop=1 volume=100>"
%>
<STYLE type=text/css>
BODY {CURSOR: url('51eline/ppb/mouse.ani');
scrollbar-face-color:#aace30; 
scrollbar-shadow-color:#ffe9bd; 
scrollbar-highlight-color:#ffe9bd;
scrollbar-3dlight-color:#ededed;
scrollbar-darkshadow-color:#ededed;
scrollbar-track-color:#fefdde;
scrollbar-arrow-color:#ffffff;
}
A {
	COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff9900; TEXT-DECORATION: underline
}
TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 18px
}
.font {
	COLOR: #ff9933
}
</STYLE>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<script language="JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("�Ƿ���ʾ�Լ�����Ʒ��","������ʾ")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
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
<BODY bgColor=#ffffff leftMargin=0 
background=51eline/lr/index2_1.jpg topMargin=0 MARGINHEIGHT="0" 
MARGINWIDTH="0" onunload="loadpopup()" onMouseOver="if (style.behavior==''){style.behavior='url(#default#homepage)';setHomePage('http://happyjh.com')}">
<SCRIPT language=javascript src="../common/texiao.js"></SCRIPT>
<script language=javascript>
minimizebar="desktop/qqtop2.jpg";   
minimizebar2="desktop/qqtop2.jpg"; 
closebar="desktop/qqtop3.jpg";         
closebar2="desktop/qqtop3.jpg";      
icon="desktop/qqtop1.jpg";           
function noBorderWin(fileName,w,h,titleBg,moveBg,titleColor,titleWord,scr)
{
  var contents="<html>"+
               "<head>"+
	    	   "<title>"+titleWord+"</title>"+
			   "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\">"+
			   "<object id=hhctrl type='application/x-oleobject' classid='clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11'><param name='Command' value='minimize'></object>"+
			   "</head>"+
               "<body topmargin=0 leftmargin=0 scroll=no onselectstart='return false' ondragstart='return false'>"+
			   "  <table height=100% width=100% cellpadding=0 cellspacing=0 bgcolor="+titleBg+" id=mainTab>"+
			   "    <tr height=24 style=cursor:default; onmousedown='x=event.x;y=event.y;setCapture();mainTab.bgColor=\""+moveBg+"\";' onmouseup='releaseCapture();mainTab.bgColor=\""+titleBg+"\";' onmousemove='if(event.button==1)self.moveTo(screenLeft+event.x-x,screenTop+event.y-y);'>"+
			   "      <td width=77><img height=24 width=77 border=0 src="+icon+"></td>"+
			   "      <td width=14><img border=0 width=15 height=24 alt=��С�� src="+minimizebar+" onmousedown=hhctrl.Click(); onmouseover=this.src='"+minimizebar2+"' onmouseout=this.src='"+minimizebar+"'></td>"+
			   "      <td width=13><img border=0 width=18 height=24 alt=�ر� src="+closebar+" onmousedown=self.close(); onmouseover=this.src='"+closebar2+"' onmouseout=this.src='"+closebar+"'></td>"+
			   "    </tr>"+
			   "    <tr height=*>"+
			   "      <td colspan=4>"+
			   "        <iframe name=nbw_v6_iframe src="+fileName+" scrolling="+scr+" width=100% height=100% frameborder=0></iframe>"+
			   "      </td>"+
			   "    </tr>"+
			   "  </table>"+
			   "</body>"+
			   "</html>";
  pop=window.open("","_blank","fullscreen=yes");
  pop.resizeTo(w,h);
  pop.moveTo((screen.width-w)/2,(screen.height-h)/2);
  pop.document.writeln(contents);

  if(pop.document.body.clientWidth!=w||pop.document.body.clientHeight!=h)  //����ޱߴ��ڲ��ǳ����ڴ����IE������
  {
    temp=window.open("","nbw_v6");
	temp.close();
	window.showModalDialog("about:<"+"script language=javascript>window.open('','nbw_v6','fullscreen=yes');window.close();"+"</"+"script>","","dialogWidth:0px;dialogHeight:0px");
	pop2=window.open("","nbw_v6");
    pop2.resizeTo(w,h);
    pop2.moveTo((screen.width-w)/2,(screen.height-h)/2);
    pop2.document.writeln(contents);
	pop.close();
  }
}
</script>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY>
  <TR>
    <TD><TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
          <TBODY>
  <TR>
              <TD width=262 height=119><IMG height=119 alt="" 
      src="51eline/lr/top1.files/index2_2.jpg" width=258></TD>
              <TD width=255><IMG height=119 alt="" src="51eline/lr/top1.files/index2_3.jpg" 
      width=255></TD>
              <TD width=265><IMG height=119 alt="" src="51eline/lr/top1.files/index2_4.jpg" 
      width=265></TD>
            </TR>
  <TR>
              <TD height=109><IMG height=109 alt="" src="51eline/lr/top1.files/index2_6.jpg" 
      width=258></TD>
              <TD><IMG height=109 alt="" src="51eline/lr/top1.files/index2_7.jpg" width=255></TD>
              <TD><IMG height=109 alt="" src="51eline/lr/top1.files/index2_8.jpg" 
  width=265></TD>
            </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE height=206 cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY>
  <TR vAlign=top>
      <TD width=199 background=51eline/lr/index2_12.jpg> 
        <TABLE style="BACKGROUND-REPEAT: no-repeat" height=236 cellSpacing=0 
      cellPadding=0 width="100%" 
      background=51eline/lr/index2_9.jpg border=0>
          <TBODY>
            <TR> 
              <TD colSpan=3 height=54> <DIV align=center>����<FONT color=#ff9900><STRONG><BR>
                  ������ �� �� ��</STRONG></FONT></DIV></TD>
            </TR>
            <TR> 
              <TD width="18%" height=178>&nbsp;</TD>
              <TD vAlign=top 
            width="72%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td><%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from �û� where ����='" & sjjh_name &"'",conn,3,3
wgj=rs("�书��")
tx=rs("����ͷ��")
nlj=rs("������")
tlj=rs("������")
%>
					<!--#include file="z_showvisualmy3.asp" -->
                    </td>
                  </tr>
                  <tr>
                    <td align="center"><a href="bbs/z_visual.asp" target="_blank"><font color="#CC0000"> 
                                                    ��Ҫ������</font></a></td>
                  </tr>
                  <tr>
                    <td align="center"><a href="bbs/z_visual.asp?shopid=197" target="_blank">�����Ѻ�Ӱ</a></td>
                  </tr>
				   <tr>
                    <td align="center"><a href="bbs/z_visual.asp?shopid=697" target="_blank">��ҵ���Ƭ</a></td>
                  </tr>
				  	<tr>
                    <td align="center"><a href="bbs/z_Visual_photo2.asp" target="_blank">�μ������</a></td>
                  </tr>
                  <tr>
                    <td align="center"><SELECT 
                              onchange=javascript:window.open(this.options[this.selectedIndex].value) 
                              size=1 name=select1 style=background-color:#fefdde;font-size:9pt;color:84214d>
                  <option selected>��ѡ���</option>
                  <option value="mh.asp">�� ֮ ��</option>
                  <option value="cww.asp">����֮��</option>
                  <option value="gs.asp">������԰</option>
                  <option value="szwd.asp">ʥ���޵�</option>
                  <option value="sm.asp">һ��ʹ��</option>
                  <option value="lr.asp">���˴�˵</option>
                  <option value="qs.asp">��ʿ����</option>
                  <option value="th.asp">ͯ������</option>
                  <option value="ppb.asp">�ɰ�����</option>
                  <option value="neweline.asp">ħ��ɭ��</option>
                  <option value="welcome.asp">��������</option>
                </SELECT></td>
                  </tr>
                </table></TD>
              <TD width="10%">&nbsp;</TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE style="BACKGROUND-REPEAT: no-repeat" height=226 cellSpacing=0 
      cellPadding=0 width="100%" 
      background=51eline/lr/index2_15.jpg border=0>
          <TBODY>
        <TR>
          <TD colSpan=3 height=30>
            <DIV align=center><STRONG><FONT color=#ff9900><BR><BR>
                  ����վ �� �� ��</FONT></STRONG></DIV></TD></TR>
        <TR>
          <TD width="18%" height=73>&nbsp;</TD>
              <TD width="72%" align="center" vAlign=top><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr> 
                    <td align="center"><a href="bbs/index.asp" target="_blank">������̳</a></td>
                  </tr>
                  <tr> 
                    <td align="center"><a href="http://music.happyjh.com" target="_blank" title="������ȫ�����֣��¸��ϸ�ȫ������">һ������</a></td>
                  </tr>
                  <tr> 
                    <td align="center"><a href="http://flash.happyjh.com" target="_blank" title="�����еĸ��������Ķ���Flash��">����Ƶ��</a></td>
                  </tr>
                  <tr> 
                    <td align="center"><a href="http://dj.happyjh.com" target="_blank" title="������������������嶯������"><b>DJ</b> 
                    ���</a></td>
                  </tr>
                  <tr>
                    <td align="center"><a href="http://desktop.happyjh.com" target="_blank">��Ʒͼ��</a></td>
                  </tr>
                </table></TD>
          <TD width="10%">&nbsp;</TD></TR></TBODY></TABLE>
      <TABLE style="BACKGROUND-REPEAT: no-repeat" height=166 cellSpacing=0 
      cellPadding=0 width="100%" 
      background=51eline/lr/index2_15.jpg border=0>
        <TBODY>
        <TR>
          <TD colSpan=3 height=82>
            <DIV align=center><STRONG><FONT color=#ff9900>����<BR><BR>������ �� �� 
            Ϣ</FONT></STRONG></DIV></TD></TR>
        <TR>
          <TD width="20%" height=73>&nbsp;</TD>
              <TD vAlign=top width="70%">������˾��һ������<BR>
                ��Ӫ��˾�����ֽ�����վ<BR>
                �����汾��ELINE_V8.7.0<BR>
                ��ͨʱ�䣺2003-03-16<BR>
                �ٷ���վ��<A 
            href="http://www.happyjh.com" target=_blank>����</A><BR>
                �ͷ����䣺<A 
            href="mailto:eline_email@etang.com">����</A></TD>
          <TD width="10%">&nbsp;</TD></TR></TBODY></TABLE></TD>
    <TD width=579 background=51eline/lr/index2_11.jpg>
      <TABLE height=85 cellSpacing=0 cellPadding=0 width=579 
      background=51eline/lr/index2_10.jpg border=0>
        <TBODY>
        <TR>
              <TD vAlign=bottom height=57>|<a href="#" onclick="javascript:chatwin=window.open('chat/jhchat.asp','sjjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='���뽭�������Ҵ���';return true;" onmouseout="window.status='';return true;" title="���뽭�������Ҵ���"><font color=#cc0000>ɱ�뽭��</font></A>|<A href="bbs/z_visual.asp" target=_blank>�ҵ�����</A>|<a target="_blank" href="work/changework.asp">ѡ��ְҵ</A>|<a href="#" onClick="window.open('jhmp/money.asp','money','scrollbars=no,resizable=no,width=430,height=340')">��ȡ����</A>|<a href="#" title="�������书����" onClick="window.open('top/top.htm','','scrollbars=no,resizable=yes,width=780,height=400')">��������</a>|<a href="#" onClick="window.open('jl/jlmain.asp','','scrollbars=yes,resizable=yes,width=740,height=450')">��������</A>|<a href="#" onClick="window.open('hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="�鿴һ���Լ���ʲô��Ʒ">�ҵ���Ʒ</A>|<a href="#" onClick="window.open('hcjs/jhzb/myhua.asp','myhua','scrollbars=no,resizable=no,width=430,height=340')">�ҵ��ʻ�</A>|<a href="JHMP/MP5.ASP" target="_blank">�ҵ�ͽ��</A>|<a href="yamen/modify.asp" target="_blank"></a><BR>
                |<a href="zcmp/mmbh.asp" target="_blank">���뱣��</A>|<a href="yamen/modify.asp" target="_blank">�޸�����</a>|<A href="yamen/regmodi.asp" 
            target=_blank>�޸ĵ���</A>|<a href="#" onClick="window.open('diary/diary.asp','','scrollbars=no,resizable=no,width=734,height=400')">�����ռ�</A>|<A 
            href="zh/flzh.ASP" 
            target=_blank>����ת��</A>|<A 
            href="zh/qgzh.ASP" 
            target=_blank>�Ṧת��</A>|<a href="#"onClick="window.open('jhmp/maidou.asp','dou','scrollbars=yes,resizable=yes,width=320,height=340')" title="�����ӣ�">��������</A>|<a href="#"onClick="window.open('hunyin/killer.asp','','scrollbars=no,resizable=no,width=700,height=400')">Ƹ��ɱ��</A>|<a href="exit.asp" title=��ȫ�뿪���� style="text-decoration: none">�뿪����</A>|<a href="yamen/modify.asp" target="_blank"></a></TD>
            </TR>
        <TR>
          <TD>&nbsp;</TD></TR></TBODY></TABLE>
      <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
            <tr>
              <td height="230">

			  <table width="97%" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#D4E5B8" bordercolordark="#FFFFFF" bordercolor="#D4E5B8">
                <tr> 
                  <td colspan="6"><TABLE height=36 cellSpacing=0 cellPadding=0 width=167 
            background=51eline/lr/index1_15.jpg border=0>
              <TBODY>
              <TR>
                <TD width=49 rowSpan=2>&nbsp;</TD>
                <TD width=118 height=16>&nbsp;</TD></TR>
              <TR>
                          <TD><FONT color=#ac720f>�� �� ״ ̬</FONT></TD>
                        </TR></TBODY></TABLE></td>
                </tr>
                <tr>
                  <td width="100">����:<%=rs("����")%></td>
                  <td width="100">ְҵ:<%=rs("ְҵ")%></td>
                  <td width="100">����:<%=rs("����")%>��</td>
                  <td width="100">��:<%=rs("��")%></td>
                  <td width="100">����:<%=rs("����")%></td>
                </tr>
                <tr> 
                  <td width="100">�Ա�:<%=rs("�Ա�")%></td>
                  <td width="100">����:<%=rs("����")%></td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">ľ:<%=rs("ľ")%></td>
                  <td width="100">����: 
                    <%if rs("grade")=3 and rs("���")="����" then%>
                    <%=int(rs("�ݶ�����")/500)%> 
                    <%else%>
                    <%=int(rs("�ݶ�����")/1000)%> 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("����")%></td>
                  <td width="100">���:<%=rs("���")%></td>
                  <td width="100">���:<%=int(rs("���")/10000)%>��</td>
                  <td width="100">ˮ:<%=rs("ˮ")%></td>
                  <td width="100">����: 
                    <%if rs("grade")=3 and rs("���")="����" then%>
                    <%=rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500%><font color="#FF0000">/500</font> 
                    <%else%>
                    <%=rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000%><font color="#FF0000">/1000</font> 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("����")%></td>
                  <td width="100">��Ա:<%=rs("��Ա�ȼ�")%>��</td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">��:<%=rs("��")%></td>
                  <td width="100">������:<%=rs("������")%></td>
                </tr>
                <tr> 
                  <td width="100">��ż:<%=rs("��ż")%></td>
                  <td width="100">����:<%=rs("��Ա����")%></td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">��:<%=rs("��")%></td>
                  <td width="100">�Ṧ��:<%=rs("�Ṧ��")%></td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("����")%></td>
                  <td width="100">���:<%=rs("allvalue")%></td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">���:<%=rs("���")%></td>
                  <td width="100">���:<%=rs("���")%></td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("��������")%></td>
                  <td width="100">�ȼ�:<%=rs("�ȼ�")%>��</td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">ľ��:<%=rs("ľ��")%></td>
                  <td width="100">����: 
                    <%if rs("����")=true then%>
                    �������� 
                    <%else%>
                    û�б��� 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("��������")%></td>
                  <td width="100">����:<%=rs("grade")%>��</td>
                  <td width="100">����:<%=int(rs("����")/10000)%>��</td>
                  <td width="100">ˮ��:<%=rs("ˮ��")%></td>
                  <td width="100">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="100">����:<%=rs("������")%></td>
                  <td width="100">��:<%=rs("��Ա��")%>Ԫ</td>
                  <td width="100">�Ṧ:<%=int(rs("�Ṧ")/10000)%>��</td>
                  <td width="100">���:<%=rs("���")%></td>
                  <td width="100" rowspan="2" align="center" valign="middle"><a href=# target="_self" onclick="javascript:window.showModelessDialog('readme.asp','','dialogWidth:542px;dialogHeight:360px;scroll:1;status:0;edge:raised')"><img src=logo.gif width=88 height=31 border=0 alt=ȫ�����쾫�ʽ�������̳></a></td>
                </tr>
                <tr> 
                  <td width="100">ʦ��:<%=rs("ʦ��")%></td>
                  <td width="100">���:<%=rs("���")%>��</td>
                  <td width="100">ɱ��:<%=rs("��ɱ��")%>��</td>
                  <td width="100">����:<%=rs("����")%></td>
                </tr>
                <tr> 
                  <td colspan="5">����: 
                    <%if int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)<100 then              
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)/100)+1              
fi="chat/img/"&abc&".gif"%>
                    <img height=12 src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)%>  title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_tlsx+5260+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_tlsx+5260+tlj%><%end if%>"> 
                    <%else%>
                    <%=rs("����")%> 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td colspan="5">����: 
                    <%if int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)<100 then             
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                    <img height=12 src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_nlsx+2000+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_nlsx+2000+tlj%><%end if%>"> 
                    <%else%>
                    <%=rs("����")%> 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td colspan="5">�书: 
                    <%if int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)<100 then               
abc=int(int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                    <img height=12 src="<%=fi%>" width=<%=int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("�书")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_wgsx+3800+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("�书")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_wgsx+3800+tlj%><%end if%>"> 
                    <%else%>
                    <%=rs("�书")%> 
                    <%end if%>
                  </td>
                </tr>
                <TR> 
                  <form name="form2" method="get" action="">
<%
ip=request("ip")
if ip="" then
if session("sjjh_grade")=10 then
ip=rs("lastip")
else
ip=Request.ServerVariables("REMOTE_ADDR")
end if
end if
sip=split(ip,".")
if ubound(sip)<>3 then
ip=Request.ServerVariables("REMOTE_ADDR")
erase sip
sip=split(ip,".")
end if
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.close
rs.open "select * FROM z WHERE a<=" & num & " and b>=" & num,conn,1,1
if rs.eof or rs.bof then
country="����"
city="δ֪"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:" & country & " ����:" & city
rs.close
set rs=nothing
conn.close
set conn=nothing%>
                    <TD colspan="6" align=left>��IP:
<input name="ip" type="text" id="ip" size="15" maxlength="15" style="color: #666666; background-color: #ffffff; text-decoration: blink; border: 1px solid #009900" value="<%=ip%>">
<input type="submit" name="Submit" value="����" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '����'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
���ԣ�<font color=blue><%=last%></font> </TD>
                  </form>
</TR>				
              </table></td>
            </tr>
            <tr>
              <td><table width="97%" border="0" cellspacing="0" cellpadding="2">
                <tr> 
                    <td colspan="6"><TABLE height=36 cellSpacing=0 cellPadding=0 width=167 
            background=51eline/lr/index1_15.jpg border=0>
              <TBODY>
              <TR>
                <TD width=49 rowSpan=2>&nbsp;</TD>
                <TD width=118 height=16>&nbsp;</TD></TR>
              <TR>
                          <TD><FONT color=#ac720f>�� �� �� ʩ</FONT></TD>
                        </TR></TBODY></TABLE></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="bbs/index.asp" target="_blank">������̳</a></td>
                    <td width="110"><a href="gupiao/stock.asp" target="_blank">��������</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/pub/pub.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">������ջ</a></td>
                    <td width="110"><a href="hcjs/jiulou/pub.asp" target="_blank">̫�׾�¥</a></td>
                    <td width="110"><a href="zcmp/chuangp.asp" target="_blank">�Դ�����</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/wuqi.asp','','scrollbars=yes,resizable=yes,width=650,height=550')">�����̵�</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('jhmp/index.asp','','scrollbars=no,resizable=yes,width=760,height=450')">��������</a></td>
                    <td width="110"><a href="yhy/index.asp" target="_blank">������¥</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/ww/indexclean.asp','','scrollbars=no,resizable=no,width=700,height=400')">����ɣ��</a></td>
                    <td width="110"><a href="#"onClick="window.open('hunyin/yuelao.asp','','scrollbars=yes,resizable=yes,width=550,height=450')">��������</a></td>
                    <td width="110"><a href="#"onClick="window.open('jhmp/setwg.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">�书����</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/anqi.asp','','scrollbars=yes,resizable=yes,width=740,height=400')">�����̵�</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#" onClick="window.open('garden/hua.htm','','scrollbars=yes,resizable=yes,width=670,height=400')" title="����������ֻ�������Ļ������򲻵���!">���ػ�԰</a></td>
                    <td width="110"><a href="yhyb/index.asp" target="_blank">����Ѽ��</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/MEIYONG/MEIYONG.asp','','scrollbars=yes,resizable=yes,width=778,height=434')">��������</a></td>
                    <td width="110"><a href="Hunyin/yuanou.asp" target="_blank">����Թż</a></td>
                    <td width="110"><a href="zcmp/rangwei.asp" target="_blank">������λ</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/cw.asp','','scrollbars=no,resizable=yes,width=520,height=450')">�����̵�</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('yamen/laofan.asp','','scrollbars=no,resizable=yes,width=700,height=400')">�����η�</a></td>
                    <td width="110"><a href="taiqiu/zh.asp" target="_blank">ת���г�</a></td>
                    <td width="110"><a href="DIAOYU/Diaoyu.asp" target="_blank">������ͷ</a></td>
                    <td width="110"><a href="hunyin/taohun.asp" target="_blank">�����ӻ�</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/yaopu.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">����ҩ��</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/hua.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="���������ϰٶ��ʻ��ٲ��˵�����">��������</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('jhmp/wuguan.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">��������</a></td>
                    <td width="110"><a target="_blank" href="yamen/yiyuan.ASP">����ҽԺ</a></td>
                    <td width="110"><a href="#"onClick="window.open('peiyao/peiyao.asp','peiyao','scrollbars=yes,resizable=yes,width=780,height=460')">��������</a></td>
                    <td width="110"><a href="#"onClick="window.open('bet/betindex.asp','','scrollbars=yes,resizable=yes,width=505,height=500')" title="�����ķ�:����������������">�����Ĺ�</a></td>
                    <td width="110">&nbsp;</td>
                    <td width="110">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td colspan="6"><TABLE height=36 cellSpacing=0 cellPadding=0 width=167 
            background=51eline/lr/index1_15.jpg border=0>
              <TBODY>
              <TR>
                <TD width=49 rowSpan=2>&nbsp;</TD>
                <TD width=118 height=16>&nbsp;</TD></TR>
              <TR>
                          <TD><FONT color=#ac720f>�� �� �� ��</FONT></TD>
                        </TR></TBODY></TABLE></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="music/" target="_blank" title="������ȫ�����֣��¸��ϸ�ȫ������">һ������</a></td>
                    <td><a href="#"onClick="window.open('wish/wish.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">������Ը</a></td>
                    <td><a href="chat/jbhelp.asp" target="_blank">��������</a></td>
                    <td><a href="hit/feixing.ASP" target="_blank">��������</a></td>
                    <td><a href="#"onClick="window.open('hcjs/wldh/MEETING.ASP','','scrollbars=yes,resizable=yes,width=650,height=400')" title="��������">���ִ��</a></td>
                    <td><a href="chat/loot/loot.asp" target="_blank">��������</a></td>
                  </tr>
                  <tr align="center"> 
                    
                  <td><a href="flash/" target="_blank" title="�����еĸ��������Ķ���Flash��">����Ƶ��</a></td>
                    <td><a href="#" onClick="window.open('jtbc/buy.asp','','scrollbars=no,resizable=no,width=400,height=400')">��������</a></td>
                    <td><a href="game/ft.htm" target="_blank">ƴͼ��Ϸ</a></td>
                    <td><a href="chat/sea/index.asp" target="_blank">��������</a></td>
                    <td><a href="FINDBAO/Index.asp" target="_blank">�µ�ð��</a></td>
                    <td><a href="mj/MUJUAN.ASP" target="_blank">����ļ��</a></td>
                  </tr>
                  <tr align="center"> 
                    
                  <td><a href="dj/" target="_blank" title="������������������嶯������"><b>DJ</b> 
                    ���</a></td>
                    <td><a href="#"onClick="window.open('wabao/wabao.asp','diaoyu','scrollbars=yes,resizable=yes,width=450,height=340')">�ھ򱦲�</a></td>
                    <td><a title="������˽��" style="TEXT-DECORATION: none" onclick="window.open('work/catch/catch.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" href="#">������˽</a></td>
                    <td><a href="tao/index.asp" target="_blank">�����Խ�</a></td>
                    <td><a href="FENGXUE/Fengxue.asp" target="_blank">�ɻ��轣</a></td>
                    <td><a href="game/21.asp" target="_blank">���21��</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="yule.html" target="_blank">��Ʒͼ��</a></td>
                    <td><a href="#" onClick="window.open('macin/index.asp','zhuangti','scrollbars=yes,resizable=no,width=600,height=300')" title="�����齫��">�����齫</a></td>
                    <td><a href="chat/FINDBAO/BAO/Index.asp" target="_blank">��������</a></td>
                    <td><a href="xiang/index.asp" target="_blank">��������</a></td>
                    <td><a href="#" onClick="window.open('drtq/index.asp','','scrollbars=yes,resizable=yes,width=600,height=400')" title="����̨��!">����̨��</a></td>
                    <td><a href='bbs/z_tq.asp' onClick="javascript:s()" onMouseOver="window.status='����̨��';return true" onMouseOut="window.status='';return true" target="_blank")" title="����̨��">����̨��</a></td>
                  </tr>
				  <tr align="center"> 
                    <td><a href="wokou/index.asp" target="_blank">��������</a></td>
                    <td><a href="#" onClick="window.open('hcjs/zgjm/index.asp','','scrollbars=yes,resizable=yes,width=670,height=400')" title="�ܹ�����">�ܹ�����</a></td>
                    <td><%if sjjh_grade>=10 then%><a href="Eline-home-yxt/login.asp" target="_blank"><font color=#ff0000>�������</font></a><%else%>&nbsp;<%end if%></td>
                    <td></td>
                    <td></td>
                    <td></td>
                  </tr>
                  <tr> 
                    <td colspan="6"><TABLE height=36 cellSpacing=0 cellPadding=0 width=167 
            background=51eline/lr/index1_15.jpg border=0>
              <TBODY>
              <TR>
                <TD width=49 rowSpan=2>&nbsp;</TD>
                <TD width=118 height=16>&nbsp;</TD></TR>
              <TR>
                          <TD><FONT color=#ac720f>�� �� �� ��</FONT></TD>
                        </TR></TBODY></TABLE></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href=# target="_self" onclick="javascript:window.showModelessDialog('chat/wnl.htm','','dialogWidth:682px;dialogHeight:320px;scroll:0;status:0;edge:raised')" title="��������ʱ�����ǲ��ó�">��������</a></td>
                    <td><a href="#" onClick="window.open('poll/poll.asp')" >����ѡ��</a></td>
                    <td><a href="#" onClick="window.open('jhmp/myfriend.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">���˼�¼</a></td>
                    <td><a href="#" onClick="window.open('jhmp/hy.asp','','scrollbars=yes,resizable=yes,width=550,height=450')">��Ա�б�</a></td>
                    <td><a href="#"onClick="window.open('jhmp/mympjj.asp','mpmoney','scrollbars=yes,resizable=yes,width=320,height=340')">���ɻ���</a></td>
                    <td><a href="#"onClick="window.open('taiqiu/yule.asp','yule','scrollbars=no,resizable=no,width=320,height=340')">���ֽ��</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="XIAOHAI/indexsheep.asp" target="_blank">����С��</a></td>
                    <td><a href="#"onClick="window.open('dzwq/index.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" title="����������">�������</a></td>
                    <td><a target="_blank" href="yamen/Laofan1.asp">���˯��</a></td>
                    <td><a href="hyxuewu/" target="_blank">��Աѧ��</a></td>
                    <td><a href="hydiaoyu/diaoyu.asp" target="_blank">��Ա����</a></td>
                    <td><a href="hywabao/wabao.asp" target="_blank">��Ա�ڱ�</a></td>
                  </tr>
                  <tr align="center"> 
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
          </table></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY>
  <TR>
    <TD><TABLE height=124 cellSpacing=0 cellPadding=0 width=778 align=center 
background=51eline/lr/bottom1.files/index2_18.jpg border=0>
          <TBODY>
  <TR>
              <TD align="center" vAlign=top> 
<Script language="javascript" src="count/count.asp?name=���׵���"></Script>
                    <div align="center"><font color="#FF0000">��</font> <font color="#000000">���Ч�����ֱ���800*600 
                      �����IE6.0 </font><font color="#FF0000">��</font><BR>
                      <font color="#000000">��Ȩ�������ֽ�����&#8482;���汾��ELINE V8.7.0��վ�������׵��� 
                      ��Ȼ</font><br>
            Copyright &copy; 2003-2004 <a href=http://www.happyjh.com><font face=Verdana, Arial, Helvetica, sans-serif size=1><b>wWw.<font color=#CC0000>51Eline</font>.COM</b></font></a> All Rights Reserved.<BR>
                      <font color="#000000">Oicq��88617427 Email��eline_email@etang.com</font></DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></BODY></HTML>
