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
<style type=text/css>
<!--
body {color:ffffff;}
td{font-size:9pt;}
a:link {color:#ffCC00;}
a:visited {color:0000ff;}
a:hover{color:#FF9900;; cursor: hand}
.hand{background-color:rgb(208,207,192);cursor:hand;}
-->
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
<BODY BGCOLOR=#006699 LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background='chat/<%=Application("aqjh_chatimage")%>' oncontextmenu=window.event.returnValue=false>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28" background='<%=Application("aqjh_chatimage")%>'>
<div align="center"><font color="#CCCCFF" size="2"><strong>�ҵ�״̬</strong></font></div>
</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="BG34.GIF">
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
              <table width="130" border="0" cellspacing="0" cellpadding="10">
                <tr>
    <td><div align="left"><img id=face src="<%=tx%>"></div></td>
  </tr>
</table>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;�Ա�:<%=rs("�Ա�")%><br>
			  &nbsp;����:<%=rs("����")%><br>
			  &nbsp;����:<%=rs("����")%><br>
              &nbsp;��ż:<%=rs("��ż")%><br>
              &nbsp;����:<%=rs("����")%><br>
			  &nbsp;����:<%=rs("������")%><br>
			  &nbsp;ʦ��:<%=rs("ʦ��")%><br>
              &nbsp;ְҵ:<%=rs("ְҵ")%> <br>
              &nbsp;����:<%=rs("����")%> <br>
              &nbsp;����:<%=rs("����")%> <br>
              &nbsp;ְλ:<%=rs("ְλ")%> <br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;���:<%=rs("���")%><br>
			  &nbsp;��Ա:<%=rs("��Ա�ȼ�")%>��<br>
			  &nbsp;����:<%=rs("��Ա����")%><br>
			  &nbsp;���:<%=rs("allvalue")%><br>
              &nbsp;�ȼ�:<%=rs("�ȼ�")%>��<br>
              &nbsp;����:<%=rs("grade")%>�� <br><font color=red>
              &nbsp;��:<%=rs("��Ա��")%>Ԫ<br>
              &nbsp;���:<%=rs("���")%>��<br>
              &nbsp;����:<%=rs("����")%>��<br>
              &nbsp;���:<%=rs("���")%>��<br></font>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("������")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("������")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;ħ��:<%=rs("ħ��")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;�ƽ�:<%=rs("�ƽ�")%><br>
              &nbsp;����:<%=rs("����")%><br>
              &nbsp;��:<%=rs("��")%><br>
	          &nbsp;֪��:<%=rs("֪��")%><br>
	          &nbsp;Ԫ��:<%=rs("Ԫ��")%><br>
              &nbsp;ɱ��:<%=rs("��ɱ��")%>��<br>
              &nbsp;ת��:<%=rs("ת��")%>��<br>
              &nbsp;����:<%=rs("����")%><br>
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
&nbsp;����:<%if rs("grade")=3 and rs("���")="����" then%>
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>                            
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000%><font color="#FF0000">/1000</font>                             
      <%end if%><br>
	  &nbsp;����:<%=rs("����")%><br>
	  &nbsp;������:<%=rs("������")%><br>
	  &nbsp;�Ṧ:<%=rs("�Ṧ")%><br>
	  &nbsp;�Ṧ��:<%=rs("�Ṧ��")%><br>
     &nbsp;���:<%=rs("���")%><font color="#FF0000">/5000000</font><br>
	 &nbsp;����:<%if rs("����")=true then%>�չ�                                        
      <%else%>
��<%end if%><br></div>
<br></table>
	  </td>
      <td width="1" bgcolor="#666699"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"></td>
    </tr></table></BODY></HTML>