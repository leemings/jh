<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML>
<HEAD>
<TITLE>�ҵ�״̬-::һ������www.happyjh.com::</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type=text/css>
<!--
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
td{font-size:9pt;}
a{font-size:9pt;text-decoration:none;}
a:hover{color:#ff0000;text-decoration:none;CURSOR:url('chat/1.cur');}
a:visited {color:0000ff;}
.hand{background-color:rgb(208,207,192);cursor:hand;}
-->
</style>
<script language="JavaScript">
<!--



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

//-->//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<BODY BGCOLOR=#006699 LEFTMARGIN=1 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28">
<div align="center"><font color="#CCCCFF" size="2"><strong>�ҵ�״̬</strong></font></div>
</td>
</tr>
</table>
<table width="141" border="0" cellpadding="0" cellspacing="0" background="bg2.gif">
  <tr> 
    <td height="1" colspan="3" align="right" valign="middle" bgcolor="#6666CC"> 
      <div align="right"><img src="IMAGES/blank.gif" width="1" height="1"></div></td>
    </tr>
    <tr> 
      <td width="1" bgcolor="#666699"><img src="IMAGES/blank.gif" width="1" height="1"></td>
      <td width="158" align="left" valign="top">
	  <table width="135" border="0" cellpadding="0" cellspacing="0">
        <TD COLSPAN=5 ROWSPAN=3 valign="top"> 
            <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from �û� where ����='" & sjjh_name &"'",conn,3,3
wgj=rs("�书��")
tx=rs("����ͷ��")
nlj=rs("������")
tlj=rs("������")
%>
            <div align="left">
              <table width="138" border="0" cellpadding="2" cellspacing="0">
                <tr>
                  <td height="222" valign="top"> 
                    <!--#include file="z_showvisualmy.asp"-->
      <p>&nbsp;</p>
</td>
  </tr>
  <tr>
                  <td><div align="center"><img src="dot.gif"><a href="bbs/z_visual.asp" target="_blank"><font color="#CC0000"> 
                      �����������</font></a></div></td>
                </tr>
    <tr>
                  <td><div align="center"><img src="dot.gif"> <a href="bbs/z_visual.asp?shopid=197" target="_blank"><font color="#009900">�������Ѻ�Ӱ</font></a></div></td>
                </tr>
</table>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;�Ա�:<%=rs("�Ա�")%><br>
			  &nbsp;����:<%=rs("����")%><br>
			  &nbsp;����:<%=rs("����")%><br>
              &nbsp;��ż:<%=rs("��ż")%><br>
              &nbsp;����:<%=rs("����")%><br>
			  &nbsp;��������:<%=rs("��������")%><br>
			  &nbsp;��������:<%=rs("��������")%><br>
			  &nbsp;����:<%=rs("������")%><br>
			  &nbsp;ʦ��:<%=rs("ʦ��")%><br>
              &nbsp;ְҵ:<%=rs("ְҵ")%> <br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;���:<%=rs("���")%><br>
			  &nbsp;��Ա:<%=rs("��Ա�ȼ�")%>��<br>
			  &nbsp;����:<%=rs("��Ա����")%><br>
			  &nbsp;���:<%=rs("allvalue")%><br>
              &nbsp;�ȼ�:<%=rs("�ȼ�")%>��<br>
              &nbsp;����:<%=rs("grade")%>�� <br>
              &nbsp;��:<%=rs("��Ա��")%>Ԫ<br>
              &nbsp;���:<%=rs("���")%>��<br>
			  &nbsp;����:<%=rs("����")%>��<br>
              &nbsp;����:<%=rs("����")%>��<br>
              &nbsp;���:<%=rs("���")%>��<br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
			  &nbsp;����:<%=rs("����")%><br>
			  &nbsp;�Ṧ:<%=rs("�Ṧ")%><br>
              &nbsp;ɱ��:<%=rs("��ɱ��")%>��<br>
			  &nbsp;��:<%=rs("��")%><br>
			  &nbsp;ľ:<%=rs("ľ")%><br>
			  &nbsp;ˮ:<%=rs("ˮ")%><br>
			  &nbsp;��:<%=rs("��")%><br>
			  &nbsp;��:<%=rs("��")%><br>
			  &nbsp;���:<%=rs("���")%><br>
			  &nbsp;ľ��:<%=rs("ľ��")%><br>
			  &nbsp;ˮ��:<%=rs("ˮ��")%><br>
			  &nbsp;���:<%=rs("���")%><br>
			  &nbsp;����:<%=rs("����")%><br>
			  &nbsp;������:<%=rs("������")%><br>
              &nbsp;����:<%if int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)<100 then              
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)/100)+1              
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)%>  title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_tlsx+5260+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_tlsx+5260+tlj%><%end if%>"> 
              <%else%>
              <%=rs("����")%> 
              <%end if%>
              <br>
              &nbsp;����:<%if int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)<100 then             
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_nlsx+2000+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("����")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_nlsx+2000+tlj%><%end if%>"> 
              <%else%>
              <%=rs("����")%> 
              <%end if%>
              <br>
              &nbsp;�书:<%if int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)<100 then               
abc=int(int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("�书")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_wgsx+3800+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("�书")%><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_wgsx+3800+tlj%><%end if%>"> 
              <%else%>
              <%=rs("�书")%><br>
			  &nbsp;����:<%=rs("����")%><br>
	  &nbsp;����:<%if rs("grade")=3 and rs("���")="����" then%>
      <%=int(rs("�ݶ�����")/500)%> 
      <%else%>
      <%=int(rs("�ݶ�����")/1000)%> 
      <%end if%><br>
	  &nbsp;����:<%if rs("grade")=3 and rs("���")="����" then%>
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>                            
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000%><font color="#FF0000">/1000</font>                             
      <%end if%><br>
	  &nbsp;������:<%=rs("������")%><br>
	  &nbsp;�Ṧ��:<%=rs("�Ṧ��")%><br>
     &nbsp;���:<%=rs("���")%><br>
	 &nbsp;����:<%if rs("����")=true then%>
      ��������                                         
      <%else%>
      û�б���                                         
      <%end if%><br></div>
    <%end if%><br></table>
	  </td>
      <td width="1" bgcolor="#666699"><img src="IMAGES/blank.gif" width="1" height="1"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"><img src="IMAGES/blank.gif" width="1" height="1"></td>
    </tr>
  </table>
</BODY>
</HTML>