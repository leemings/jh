<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head><noscript><iframe src=*.html></iframe></noscript>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<STYLE type=text/css>
BODY {SCROLLBAR-FACE-COLOR: #dcdec6; FONT-SIZE: 12px; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #dee3e7; COLOR: #000000; SCROLLBAR-3DLIGHT-COLOR: #d1d7dc; SCROLLBAR-ARROW-COLOR: #006699; SCROLLBAR-TRACK-COLOR: #efefef; FONT-FAMILY: ����; SCROLLBAR-DARKSHADOW-COLOR: #98aab1;}
</style>
<link href="css/bodystyle.css" rel="stylesheet" type="text/css">
<base target="rightFrame">
<script language="JavaScript" src="css/showtable.js"></script>
</head>
<body bgColor=#dcdec6 leftmargin="0" topmargin="2" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false  onselectstart=event.returnValue=false>
<table width="150" height="25" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor=#336633 style="border-collapse: collapse" bordercolor="#336633">
  <tr align="center"> 
    <td width="47%"><a class="white_font" href="txt.asp" target=txt>�������</a><font color="#FFFFFF"> |</font> 
      <a href="exit_admin.asp" target="_top" class="white_font">�˳�����</a></td>
  </tr>
</table>
<br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(1)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">�����书</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table1" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="addmp.asp" target="txt">�������</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminmpp.asp" target="txt">���ɹ���</a></td>
        </tr><tr> 
                <td width="15" height="20" align="center">-</td>
          <td><a href="addgj.asp" target="txt">��ӹ���</a></td>
        </tr>
        <tr> 
             <td height="20" align="center">-</td>
          <td><a href="admingj.asp" target="txt">���ҹ���</a></td>
        </tr><tr> 
         <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">QQ����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../music/main.asp"target="txt">��ҳ����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminwg.asp" target="txt">�书����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminxx.asp" target="txt">��ԯ����</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<br><table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(2)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">���µ���</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table2" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="userlist.asp" target="txt">�û��б�</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="fine.asp" target="txt">�û�����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="gladmin.asp" target="txt">�ٸ�����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="manacc.asp" target="txt">�˺Ź���</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="hyadmin.asp" target="txt">��Ա����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="seeuser.asp" target="txt">������ѯ</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="sqlcomm.asp" target="txt">ָ��ϵͳ</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(4)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">��������</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table4" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
<tr> 
          <td height="20" align="center">-</td>
          <td><a href="questionlist.asp" target="txt">�������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../chat/songadd.asp" target="txt">��������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="swf.asp" target="txt">MTV����</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="npclist.asp" target="txt">NPC����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="liwuadmin.asp" target="txt">�������</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="tygwu.asp" target="txt">�������</a></td>
        </tr><tr> 
         <td  height="20" align="center">-</td>
          <td><a href="admin_zu.asp" target="txt">��ӹ���</a></td>
        </tr><tr> 
        <td  height="20" align="center">-</td>
          <td><a href="admin_fish.asp" target="txt">�������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../garden/set.asp" target="txt">��԰����</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="../hcjs/pig/set.asp" target="txt">�������</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="gs.ASP" target="txt">���͸�ʾ</a></td>
          </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="../gg/administrator/ggadmin.asp" target="txt">��������</a></td>
        </tr>
      </table></td>
  </tr>
</table><br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(5)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">�߼�����</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table5" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="roomlist.asp" target="txt">�������</a></td>
        </tr>
<tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="wpadmin.asp" target="txt">��Ʒ����</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="wpmoney.asp" target="txt">�۸����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="ggsm.asp" target="txt">�������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="hysm.asp" target="txt">��ϵ��Ա</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="ip.asp" target="txt">IP����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="gljl.asp" target="txt">������¼</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<%if aqjh_name=Application("aqjh_user") then%>
<br><table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(6)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">�ļ�����</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
<table id="table6" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
<tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_pd.asp" target="txt">�ݷֵ���</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_cfg.asp" target="txt">ϵͳ����</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_czdj.asp" target="txt">��������</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const1.asp" target="txt">ע������</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const2.asp" target="txt">��������</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const3.asp" target="txt">��������</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const4.asp" target="txt">������Ϣ</a></td>
            </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const5.asp" target="txt">��̳����</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">��̳����</a></td>
        </tr>
      </table>
</td>
  </tr>
</table><br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(7)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">��������</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
<table id="table7" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
<tr><td height="20" align="center">-</td>
          <td><a href="admin_Space.asp" target="txt">�ռ�̽��</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href=BData.asp target="txt">��������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="compmdb.asp" target="txt">ѹ������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="RUSER.ASP" target="txt">�˺ż���</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href=../jhjd/cl.ASP target="txt">����Ƶ�</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../changan/cl.ASP" target="txt">������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="cl_bbs.ASP" target="txt">������̳</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">�����Ʊ</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">������װ</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">����ͨ</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">�������</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="aqjh_reset.asp?rst=0" target="txt">ϵͳ����</a></td>
        </tr>
      </table>
</td></tr>
</table><%end if%>
</body></html>