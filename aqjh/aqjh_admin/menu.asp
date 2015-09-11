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
<title>后台管理系统</title>
<STYLE type=text/css>
BODY {SCROLLBAR-FACE-COLOR: #dcdec6; FONT-SIZE: 12px; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #dee3e7; COLOR: #000000; SCROLLBAR-3DLIGHT-COLOR: #d1d7dc; SCROLLBAR-ARROW-COLOR: #006699; SCROLLBAR-TRACK-COLOR: #efefef; FONT-FAMILY: 宋体; SCROLLBAR-DARKSHADOW-COLOR: #98aab1;}
</style>
<link href="css/bodystyle.css" rel="stylesheet" type="text/css">
<base target="rightFrame">
<script language="JavaScript" src="css/showtable.js"></script>
</head>
<body bgColor=#dcdec6 leftmargin="0" topmargin="2" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false  onselectstart=event.returnValue=false>
<table width="150" height="25" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor=#336633 style="border-collapse: collapse" bordercolor="#336633">
  <tr align="center"> 
    <td width="47%"><a class="white_font" href="txt.asp" target=txt>管理帮助</a><font color="#FFFFFF"> |</font> 
      <a href="exit_admin.asp" target="_top" class="white_font">退出管理</a></td>
  </tr>
</table>
<br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(1)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">门派武功</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table1" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="addmp.asp" target="txt">添加门派</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminmpp.asp" target="txt">门派管理</a></td>
        </tr><tr> 
                <td width="15" height="20" align="center">-</td>
          <td><a href="addgj.asp" target="txt">添加国家</a></td>
        </tr>
        <tr> 
             <td height="20" align="center">-</td>
          <td><a href="admingj.asp" target="txt">国家管理</a></td>
        </tr><tr> 
         <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">QQ管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../music/main.asp"target="txt">首页歌曲</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminwg.asp" target="txt">武功管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="adminxx.asp" target="txt">轩辕管理</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<br><table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(2)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">人事档案</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table2" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="userlist.asp" target="txt">用户列表</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="fine.asp" target="txt">用户管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="gladmin.asp" target="txt">官府设置</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="manacc.asp" target="txt">账号管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="hyadmin.asp" target="txt">会员管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="seeuser.asp" target="txt">条件查询</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="sqlcomm.asp" target="txt">指令系统</a></td>
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
          <td valign="middle"><font color="#FFFFFF">基本管理</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table4" width="150" border="0" align="center" cellpadding="0" cellspacing="2">
<tr> 
          <td height="20" align="center">-</td>
          <td><a href="questionlist.asp" target="txt">字迷题库</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../chat/songadd.asp" target="txt">歌曲管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="swf.asp" target="txt">MTV管理</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="npclist.asp" target="txt">NPC管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="liwuadmin.asp" target="txt">礼物管理</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="tygwu.asp" target="txt">怪物管理</a></td>
        </tr><tr> 
         <td  height="20" align="center">-</td>
          <td><a href="admin_zu.asp" target="txt">组队管理</a></td>
        </tr><tr> 
        <td  height="20" align="center">-</td>
          <td><a href="admin_fish.asp" target="txt">钓鱼管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../garden/set.asp" target="txt">花园大赛</a></td>
        </tr><tr> 
        <td height="20" align="center">-</td>
          <td><a href="../hcjs/pig/set.asp" target="txt">养猪大赛</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="gs.ASP" target="txt">发送告示</a></td>
          </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="../gg/administrator/ggadmin.asp" target="txt">发布公告</a></td>
        </tr>
      </table></td>
  </tr>
</table><br>
<table width="150" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#336633">
  <tr> 
    <td height="25" bgcolor=#336633><table width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="cursor:hand;" onclick="showtable(5)">
        <tr> 
          <td width="15" align="center" valign="middle"></td>
          <td valign="middle"><font color="#FFFFFF">高级管理</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table id="table5" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
        <tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="roomlist.asp" target="txt">房间管理</a></td>
        </tr>
<tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="wpadmin.asp" target="txt">物品管理</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="wpmoney.asp" target="txt">价格调整</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="ggsm.asp" target="txt">广告设置</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="hysm.asp" target="txt">联系会员</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href="ip.asp" target="txt">IP管理</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="gljl.asp" target="txt">操作记录</a></td>
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
          <td valign="middle"><font color="#FFFFFF">文件配置</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
<table id="table6" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
<tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_pd.asp" target="txt">泡分调整</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_cfg.asp" target="txt">系统配置</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_czdj.asp" target="txt">条件限制</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const1.asp" target="txt">注册配置</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const2.asp" target="txt">文字设置</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const3.asp" target="txt">开关设置</a></td>
        </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const4.asp" target="txt">进入信息</a></td>
            </tr><tr> 
          <td width="15" height="20" align="center">-</td>
          <td><a href="sys_const5.asp" target="txt">论坛奖罚</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">论坛管理</a></td>
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
          <td valign="middle"><font color="#FFFFFF">数据整理</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
<table id="table7" width="150" border="0" align="center" cellpadding="0" cellspacing="2" style="display:none;">
<tr><td height="20" align="center">-</td>
          <td><a href="admin_Space.asp" target="txt">空间探测</a></td>
        </tr>
        <tr> 
          <td height="20" align="center">-</td>
          <td><a href=BData.asp target="txt">备份数据</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="compmdb.asp" target="txt">压缩数据</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="RUSER.ASP" target="txt">账号急救</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href=../jhjd/cl.ASP target="txt">清理酒店</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="../changan/cl.ASP" target="txt">清理房屋</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="cl_bbs.ASP" target="txt">清理论坛</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">清理储物</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">清理股票</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">清理秀装</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">清理交通</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="bbsgl.asp" target="txt">清理贷款</a></td>
        </tr><tr> 
          <td height="20" align="center">-</td>
          <td><a href="aqjh_reset.asp?rst=0" target="txt">系统重启</a></td>
        </tr>
      </table>
</td></tr>
</table><%end if%>
</body></html>