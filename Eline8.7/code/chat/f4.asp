<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>��һ�������wWw.happyjh.com��</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=javascript>
minimizebar='ico';
closebar="close.gif"; 
icon="icon.gif"; 
function noBorderWin(fileName,w,h,titleBg,titleColor,titleWord,borderColor,scr)
{
  nbw=window.open('','','fullscreen=yes');
  nbw.resizeTo(w,h);
  nbw.moveTo((screen.width-w)/2,(screen.height-h)/2);
  nbw.document.writeln("<title>"+titleWord+"</title>");
  nbw.document.writeln("<"+"script language=javascript"+">"+"cv=0"+"<"+"/script"+">");
  nbw.document.writeln("<body topmargin=0 leftmargin=0 scroll=no>");
  nbw.document.writeln("<table cellpadding=0 cellspacing=0 width=100% height=100% style=\"border: 1px solid "+borderColor+"\">");
  nbw.document.writeln("  <tr bgcolor="+titleBg+" height=20 onselectstart='return false' onmousedown='cv=1;x=event.x;y=event.y;setCapture();' onmouseup='cv=0;releaseCapture();' onmousemove='if(cv)window.moveTo(screenLeft+event.x-x,screenTop+event.y-y);'>");
  nbw.document.writeln("    <td style=\"font-family:����; font-size:12px; color:"+titleColor+";\"><span style='top:1;position:relative;cursor:default;'>"+titleWord+"</span></td>");
  nbw.document.writeln("    <td align=right width=30 style=font-size:16px;>");
  nbw.document.writeln("      <img border=0 width=12 height=12 alt=��С�� src="+minimizebar+" onclick=window.blur();>");
  nbw.document.writeln("      <img border=0 width=12 height=12 alt=�ر� src="+closebar+" onclick=window.close();>");
  nbw.document.writeln("    </td>");
  nbw.document.writeln("  </tr>");
  nbw.document.writeln("  <tr>");
  nbw.document.writeln("    <td colspan=3><iframe frameborder=0 width=100% height=100% scrolling="+scr+" src="+fileName+"></iframe></td>");
  nbw.document.writeln("  </tr>");
  nbw.document.writeln("</table>");
}
</SCRIPT>

<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;CURSOR:url('1.cur');}
</style>
<script Language="Javascript">
function s(){parent.f2.document.af.mdsx.checked=false;parent.f2.document.af.sytemp.focus();}
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
function showname(){
parent.f2.document.af.mdsx.checked=true;parent.f2.document.af.sytemp.focus();
if(parent.m.location.href=="about:blank"){
parent.m.location.href="f3.asp";}
else
{parent.m.location.reload();}
}
function ex(){parent.t.location.href='about:blank';top.location.href="exitlt.asp";return true;}
</script>

</head>

<body background="f2.gif" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td valign="top" background="]">
	<table width="160" height="20" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#E1D7A2" bordercolordark="#efefef" background="yinyue/bg2.gif" bgcolor="#efefef">
        <tr>
          <td valign="top"> 
            <div align="center"> 
              <IFRAME src="yinyue/bgm.htm" width="150" height="15" scrolling="no" frameborder="0"></iframe>
            </div></td>
  </tr>
</table>
	  <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="f2.gif">
        <tr align="center"> 
          <td> <a href="selectchatroom.asp" onClick="javascript:s()" onMouseOver="window.status='��������';return true" onMouseOut="window.status='';return true" target="f3" title="ѡ�񷿼�">����</font></a></td>
          <td><a href="#" onMouseOver="window.status='����������������������';return true" onMouseOut="window.status='';return true" onClick="javascript:showname();" title="ˢ��������������">����</a></td>
          <td><a href="../CHANGAN/xiaowu.asp" onClick="javascript:s()" target="_blank" onMouseOver="window.status='����С��';return true" onMouseOut="window.status='';return true" title="����С��">С��</a></td>
          <td><a href="npc/a1.asp" onclick="javascript:s()" title="���߹���" target="f3">����</a></td>
        </tr>
        <tr align="center"> 
          <td> <a href="caidan.asp" onClick="javascript:s()" onMouseOver="window.status='���ܿ�';return true" onMouseOut="window.status='';return true" target="f3" title="���ܿ�"><font color="#00FFFF">�˵�</font></a></td>
          <td><a href="KiNg.asp" onClick="javascript:s()" onMouseOver="window.status='��������ʹ��';return true" onMouseOut="window.status='';return true" target="f3" title="������ʩ�볣�ù���">����</font></A></td>
          <td><a href='command.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='������';return true" onMouseOut="window.status='';return true" title="��һ�¶���">����</font></a> 
          </td>
          <td> <a href="cw/cw.asp" onClick="javascript:s()" target=f3 title="�չ��Լ���С����">����</font></a></td>
        </tr>
        <tr align="center"> 
          <td><a href="savevalue.asp" onClick="javascript:s()" onMouseOver="window.status='�������澭��ֵ������ʾ��ǰ״̬��';return true" onMouseOut="window.status='';return true" target="f3" title="�鿴��ǰ���">���</font></a></td>
          <td><a href="game/game_index.htm" onclick="javascript:s()" title="������ѡ����ʮ�����С��Ϸ" target="f3">��Ϸ</font></a> 
            <a href='setfontsize.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='�ı�Ի�������İ�ֵ���о࣬������ɺ����㡰���������������á�';return true" onMouseOut="window.status='';return true" title="�ı������С���о�"></a></font></td>
          <td><a href="BOY/Boy.asp" onClick="javascript:s()" target="f3" title="ϸ���չ�����С��">С��</font></a> 
          </td>
          <td> <a href="#" onClick="window.open('../hcjs/jhjs/vehicle.asp','','scrollbars=yes,resizable=yes,width=670,height=400')" title="����ͨ����">��ͨ</font></a></td>
        </tr>
        <tr align="center"> 
          <td height="20"><a href="wupin.asp" onClick="javascript:s()" onMouseOver="window.status='��������������Щ��Ʒ';return true" onMouseOut="window.status='';return true" target="f3" title="��������ʲô��Ʒ">��Ʒ</font></a></td>
          <td height="20"><a href='setwg3.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='���С�';return true" onMouseOut="window.status='';return true" title="���ɵ��书">����</font></a></td>
          <td height="20"><a href='pic.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='�ı�Ի�������İ�ֵ���о࣬������ɺ����㡰���������������á�';return true" onMouseOut="window.status='';return true" title="�����">ͼ��</a></td>
          <td height="20"><a href="xlwupin.asp" onClick="javascript:s()" onMouseOver="window.status='������Ʒ��';return true" onMouseOut="window.status='';return true" target="f3" title="��������">����</font></a></td>
        </tr>
        <tr align="center"> 
          <td><a href="#" onClick="javascript:parent.qp()" onMouseOver="window.status='����Ի��������жԻ���';return true" onMouseOut="window.status='';return true" title="�뼰ʱ��������¼">����</font></a></td>
          <td><a href='killman.asp' onClick="javascript:s()" onMouseOver="window.status='�鿴�������ˡ�';return true" onMouseOut="window.status='';return true" target="f3")" title="�鿴��������">����</font></a></td>
          <td><a href="song.asp" onMouseOver="window.status='�����Լ����������ѡ�';return true" onMouseOut="window.status='';return true" target="f3" title="�������">���</font></a></td>
          <td><a href='../10top.asp' onClick="javascript:s()" onMouseOver="window.status='�鿴����10��';return true" onMouseOut="window.status='';return true" target="f3")" title="�鿴�������а�">����</font></a></td>
        </tr>
		<tr align="center"> 
          <td><a href="../tq/Weather.asp" title="ȫ������Ԥ��" target="f3" onClick="javascript:s()" onMouseOver="window.status='ȫ������Ԥ��';return true" onMouseOut="window.status='';return true">����</font></a></td>
          <td><a href='../garden/hua.asp' onClick="javascript:s()" onMouseOver="window.status='���������������';return true" onMouseOut="window.status='';return true" target="_blank")" title="���������������">��԰</font></a></td>
          <td><a href="#" onClick="window.open('../drtq/index.asp','','scrollbars=yes,resizable=yes,width=600,height=400')" title="����̨��!">̨��</a></td>
          <td><a href='st.asp' onClick="javascript:s()" onMouseOver="window.status='���߹㲥�͵���';return true" onMouseOut="window.status='';return true" target="f3")" title="���ֲ��ϣ��������ʣ�">����</a></td>
        </tr>
		<tr align="center"> 
          <td><a href="../bbs/index.asp" title="��������̳��վ�����ڽ����Ϳ϶�������" target="_blank" onClick="javascript:s()" onMouseOver="window.status='��������̳��վ�����ڽ����Ϳ϶�������';return true" onMouseOut="window.status='';return true">��̳</a></td>
          <td><a href='setwg4.asp' onClick="javascript:s()" onMouseOver="window.status='������ԯ�书';return true" onMouseOut="window.status='';return true" target="f3")" title="������ԯ�书">��ԯ</font></a></td>
          <td><a href='empdz/mp.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='���ɴ�ս';return true" onMouseOut="window.status='';return true" title="���ɴ�ս">�۽�</font></a></td>
          <td><a href='SETWGdb.ASP' onClick="javascript:s()" onMouseOver="window.status='���ɴ�ս�����Ὥ������';return true" onMouseOut="window.status='';return true" target="f3")" title="�ᱦ����ʱר�õ��书">�ᱦ</a></td>
        </tr>
        <tr align="center"> 
          <td><a href="../zhuangtai.asp" onClick="javascript:s()" target="f3" title="�ҵ�״̬"><font color="#FFCCFF">״</font></a><a href="#" title="�鿴�Լ�����ϸ״̬" onMouseOver="window.status='�鿴�Լ�����ϸ״̬��';return true" onMouseOut="window.status='';return true" onclick="javascript:window.open('zt/index.asp','zhuantai','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=500,height=530')" target="_self"><font color="#99CCFF">̬</font></a> 
          </td>
          <td><a href="stuntlist.asp" onClick="javascript:s()" target="f3" title="ʹ���ؼ�">�ؼ�</a></td>
          <td><a href="face.asp" onClick="javascript:s()" target="f3" title="ѡ����ϲ����ͷ��">ͷ��</a></td>
          <td><a href="#" onMouseOver="window.status='�˳������ң����Զ����澭��ֵ��';return true" onMouseOut="window.status='';return true" onclick="javascript:ex()" title="��ȫ�뿪�����">�뿪</font></a></td>
        </tr>
		</table>   
</td>    
  </tr>      
</table>   
</body>   
</html>   
   
   
   
   
   
   
  
  
  
  
  
  
