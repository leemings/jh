<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML>
<HEAD>
<TITLE>�ҵ�״̬</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
chatimage=Application("aqjh_chatimage")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type="text/css">BODY { COLOR: #000000; FONT-FAMILY: "����"; FONT-SIZE: 9pt } TD { COLOR: #000000; 
FONT-FAMILY: "����"; FONT-SIZE: 9pt } A:link { COLOR: #000000;
FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none } A:visited { COLOR: 
#000000;FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: 
none } A:hover { COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; 
TEXT-DECORATION: underline blink } SELECT { BACKGROUND-COLOR: #ffcc00; BORDER-BOTTOM: 
#ffcc66 solid; BORDER-LEFT: #999900 solid; BORDER-RIGHT: #ffcc66 solid; BORDER-TOP: 
#999900 solid; FONT-SIZE: 9pt } INPUT { BACKGROUND-COLOR: RED; BORDER-BOTTOM: 
1px groove; BORDER-LEFT: 1px groove; BORDER-RIGHT: 1px groove; BORDER-TOP: 1px 
groove; COLOR: #ffffff; FONT-SIZE: 9pt } TD { }
BODY {
SCROLLBAR-FACE-COLOR: white;
 SCROLLBAR-HIGHLIGHT-COLOR: white;
 SCROLLBAR-SHADOW-COLOR: white;
 SCROLLBAR-3DLIGHT-COLOR: white;
 SCROLLBAR-ARROW-COLOR: white;
 SCROLLBAR-TRACK-COLOR: white;
 SCROLLBAR-DARKSHADOW-COLOR: white;
 SCROLLBAR-BASE-COLOR: white
} 
</style>
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
<BODY BGCOLOR=white LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background='chat/<%=Application("aqjh_chatimage")%>' oncontextmenu=window.event.returnValue=false>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="##FF3366">
  <tr>
<td width="100%" height="28" background='<%=Application("aqjh_chatimage")%>'>
<div align="center"><font color="#FF6600" size="2"><strong>��������</strong></font></div>
</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/bg.jpg">
  <tr> 
    <td height="1" colspan="3" align="right" valign="middle" bgcolor="#6666CC"> </td>
    </tr>
    <tr> 
      <td width="1" bgcolor="#666699"> </td>
      <td width="158" align="left" valign="top">
	  <table width="135" border="0" cellpadding="0" cellspacing="0">
        <TD COLSPAN=5 ROWSPAN=3 valign="top"> 
            <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from �û� where ����='" & aqjh_name &"'",conn,3,3
wgj=rs("�书��")
tx=rs("����ͷ��")
nlj=rs("������")
tlj=rs("������")
%>
            <div align="left">
              <table width="500" border="0" cellspacing="0" cellpadding="10">
                <tr>
    <td><div align="left"><img id=face src="<%=tx%>"></div></td>
  </tr>
</table>
              <font color=red>&nbsp;����:<%=rs("����")%>&nbsp;&nbsp;&nbsp;�Ա�:<%=rs("�Ա�")%>
			  &nbsp;&nbsp;&nbsp;����:<%=rs("����")%>
			  &nbsp;&nbsp;&nbsp;����:<%=rs("����")%>
              &nbsp;&nbsp;&nbsp;��ż:<%=rs("��ż")%>
              &nbsp;&nbsp;&nbsp;����:<%=rs("����")%><br><br>
			  &nbsp;����:<%=rs("������")%>
			  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʦ��:<%=rs("ʦ��")%>
              &nbsp;&nbsp;&nbsp;ְҵ:<%=rs("ְҵ")%> 
              &nbsp;����:<%=rs("����")%>
              &nbsp;&nbsp;&nbsp;���:<%=rs("���")%><br><br>
			  &nbsp;��Ա:<%=rs("��Ա�ȼ�")%>��
			  &nbsp;&nbsp;&nbsp;����:<%=rs("��Ա����")%>
			  &nbsp;&nbsp;&nbsp;���:<%=rs("allvalue")%>
              &nbsp;&nbsp;&nbsp;�ȼ�:<%=rs("�ȼ�")%>��
              &nbsp;&nbsp;&nbsp;����:<%=rs("grade")%>�� <br><br>
              &nbsp;��:<%=rs("��Ա��")%>Ԫ
              &nbsp;&nbsp;&nbsp;���:<%=rs("���")%>��
              &nbsp;&nbsp;&nbsp;����:<%=rs("����")%>��
              &nbsp;&nbsp;&nbsp;���:<%=rs("���")%>��<br><br>
              &nbsp;����:<%=rs("����")%>
              &nbsp;&nbsp;&nbsp;����:<%=rs("������")%>
              &nbsp;&nbsp;&nbsp;����:<%=rs("����")%>
              &nbsp;&nbsp;&nbsp;����:<%=rs("������")%>
              &nbsp;&nbsp;&nbsp;����:<%=rs("����")%><br><br>
              &nbsp;����:<%=rs("����")%>
	      &nbsp;&nbsp;&nbsp;֪��:<%=rs("֪��")%>
              &nbsp;&nbsp;&nbsp;ɱ��:<%=rs("��ɱ��")%>��
              &nbsp;&nbsp;&nbsp;ת��:<%=rs("ת��")%>��
              &nbsp;����:<%=rs("����")%>
              &nbsp;״Ԫ:<%=rs("״Ԫ")%><br><br>
			  &nbsp;��:<%=rs("��")%>
			  &nbsp;&nbsp;&nbsp;ľ:<%=rs("ľ")%>
			  &nbsp;&nbsp;&nbsp;ˮ:<%=rs("ˮ")%>
			  &nbsp;&nbsp;&nbsp;��:<%=rs("��")%>
			  &nbsp;&nbsp;&nbsp;��:<%=rs("��")%><br><br>
			  &nbsp;���:<%=rs("���")%>
			  &nbsp;&nbsp;&nbsp;ľ��:<%=rs("ľ��")%>
			  &nbsp;&nbsp;&nbsp;ˮ��:<%=rs("ˮ��")%>
			  &nbsp;&nbsp;&nbsp;���:<%=rs("���")%>
			  &nbsp;&nbsp;&nbsp;����:<%=rs("����")%>
			  &nbsp;������:<%=rs("������")%><br><br>
&nbsp;����:<%if rs("grade")=3 and rs("���")="����" then%>
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500%>/500 
      <%else%>                            
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000%>/1000                            
      <%end if%>
	&nbsp;&nbsp;&nbsp;����:<%=rs("����")%>
	&nbsp;&nbsp;&nbsp;������:<%=rs("������")%>
	&nbsp;&nbsp;&nbsp;�Ṧ:<%=rs("�Ṧ")%><br><br>
	&nbsp;&nbsp;&nbsp;�Ṧ��:<%=rs("�Ṧ��")%>
     &nbsp;���:<%=rs("���")%>/5000000
	&nbsp;&nbsp;&nbsp;����:<%if rs("����")=true then%>�չ�                                        
      <%else%>
��<%end if%><br></font></div>
<br></table>
	  </td>
      <td width="1" bgcolor="#666699"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"></td>
    </tr></table></BODY></HTML>
