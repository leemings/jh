<%
if Session("aqjh_name")="" then 
Response.Write "<script Language=Javascript>alert('��û�е�½�����������Ѿ��Ͽ�����!');window.close();</script>"
response.end
end if%>
<script lanaguage=javascript>
//if(window.name!="aqjh_win"){ var i=1;while (i<=50){window.alert("������ʲôѽ���㵹��ˢ��������������50�Σ���");i=i+1;}top.location.href="../../exit.asp"}
</script>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>-��������</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY background=../../bg.gif leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<center><br><font color=green size=6 face="����">�� �� �� ��</font><br><br><br>
<table width="80%" border="1" cellspacing="2" cellpadding="2" bordercolor="#CC9933" align="center" style=font-size:9pt>
  <tr> 
    <td colspan="3"> 
      <div align="center"><img src="badman.gif" width="179" height="150"></div>
    </td>
  </tr>
  <tr> 
    <td rowspan="3" width="33%"><b>������������������Ϸ��غ�����Щ�ձ�������ɧ�����ǣ���Ϊһ����������ͣ����ǸϿ����<font color="#FF0000">��������</font>��ս��������</b> 
    </td>
    <td width="40%" height="21"><font color="#996633"><b><font color="#CC0033">���������Ϳ�������ս������</font></b></font></td>
    <td width="27%" height="21"><a href="wokou1.asp" target="_self">̤��ս��</a></td>
  </tr>
  <tr> 
    <td width="40%" height="23"><font color="#996633"><b><font color="#CC0033">���м����Ϳ�������ս������</font></b></font></td>
    <td width="27%" height="23"><a href="wokou2.asp" target="_self">̤��ս��</a></td>
  </tr>
  <tr> 
    <td width="40%"><font color="#996633"><b><font color="#CC0033">���߼����Ϳ�������ս������</font></b></font></td>
    <td width="27%"><a href="wokou3.asp" target="_self">̤��ս��</a></td>
  </tr>
</table>
<p align=center><font style=font-size:9pt color=#000000>��Ȩ���С����ֽ�����վ��</font></p>
</BODY></HTML>