<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("hxf_c_chatroombgimage")
chatroombgcolor=Application("hxf_c_chatroombgcolor")
nl=""
Select Case id
Case "000"
 nl="程序所在目录不是虚拟目录，或设置有错误，Global.asa 没有被执行。本程序需要虚拟目录的支持！"
Case "100"
 nl="对不起，你尚未登录或已经超时断开连接，请回到登录页面重新输入用户名和密码进行登录。"
Case "101"
 nl="必须拥有 3 级以上权限才能添加自创的动作。"
Case "102"
 nl="各个项目均不能放空白，请填写完整！"
Case "103"
 nl="动作内容必须是以“//”开头的语句！"
Case "104"
 nl="动作内容不能出现“//##”，这将导致发言后姓名出现重复，请将 ## 去掉！"
Case "105"
 nl="动作名称或动作内容超过限制的长度！"
Case "106"
 nl="动作内容含有“%%”，动作类型应改为“1 对别人”！"
Case "107"
 nl="动作类型为“1 对别人”，动作内容中却没有出现“%%”！"
Case "108"
 nl="添加的动作内容中，不能包含半角的“\”、“｜”、“" & chr(34) & "”！"
Case "109"
 nl="该内容已经存在，不必重复添加。"
Case "120"
 nl="你没有清理“聊务公开”的权限。"
Case "121"
 nl="没有超过 7 天的记录，不能清除。"
Case "130"
 nl="你没有“帐号管理”的权限。"
Case "131"
 nl="没有此类帐号可供删除。"
Case "132"
 nl="请输入用户名。"
Case "133"
 nl="在已“自杀”的帐号中没有找到该用户名，不能恢复。"
Case "134"
 nl="不能恢复用户名：<font color=red>" & Request.QueryString("name") & "</font>，因为系统中已有相同的用户名存在。如果你确实想要恢复该用户名，则请用“删除帐号”功能，先删除系统中相同的用户名后再恢复该用户名。"
Case "135"
 nl="请输入欲删除的用户名。"
Case "136"
 nl="用户名不存在，不能删除。"
Case "137"
 nl="用户名不存在。"
Case "138"
 nl="该用户名本已永久保留，不必重复操作。"
Case "139"
 nl="该用户名未被永久保留，不能取消。"
Case "140"
 nl="旧用户名和新用户名均不能为空。"
Case "141"
 nl="不能将用户名改为：<font color=red>" & Request.QueryString("name") & "</font>，因为系统中已有相同的用户名存在。如果你确实想要改成这个用户名，则请用“删除帐号”功能，先删除系统中相同的用户名。"
Case "142"
 nl="用户名和新密码均不能为空。"
Case "150"
 nl="你没有“数据压缩”的权限。"
Case "151"
 nl="数据库尚未关闭，不必重新打开数据库。"
Case "152"
 nl="数据库尚未打开，不必关闭数据库。"
Case "160"
 nl="请输入要搜索的用户名。"
Case "200"
 nl="你没有“踢人”的权限。"
Case "201"
 nl="你不在聊天室中，不能执行“踢人”操作。"
Case "202"
 nl="请指定要踢出的对象。"
Case "203"
 nl="不得无故踢人，请输入原因。"
Case "204"
 nl="用户名：<font color=red>" & Request.QueryString("kickname") & "</font> 不在聊天室中，不必再踢了。"
Case "205"
 nl="呵呵，毛病，踢自己玩啊。"
Case "210"
 nl="你没有“IP管理”的权限。"
Case "211"
 nl="你不在聊天室中，不能进行“IP管理”。"
Case "212"
 nl="请指定要封锁的 IP。"
Case "213"
 nl="封锁自己的 IP？别傻了！"
Case "214"
 nl="不得无故封锁 IP，请输入原因。"
Case "215"
 nl="该 IP 已经被封锁了，不能重复封锁。"
Case "216"
 nl="请输入解锁该 IP 的原因。"
Case "217"
 nl="该 IP 未被封锁，不能进行解锁。"
Case "218"
 nl="请指定要解锁的 IP。"
Case "219"
 nl="要封锁的IP与用户名不对应。"
Case "220"
 nl="你没有“炸弹操作”的权限。"
Case "221"
 nl="你不在聊天室中，不能进行“炸弹操作”。"
Case "222"
 nl="请指定要轰炸的对象。"
Case "223"
 nl="啊，你不是厕所里打灯笼——找死吧。"
Case "224"
 nl="不得无故乱扔炸弹，请输入原因。"
Case "225"
 nl="用户名：<font color=red>" & Request.QueryString("bombname") & "</font> 不在聊天室中，炸不到了。"
Case "230"
 nl="你没有更改“系统参数”的权限。"
Case "231"
 nl="新值与旧值完全相同，不必进行更改。"
Case "232"
 nl="输入的值不合法。"
Case "240"
 nl="你没有进行“级别管理”的权限。"
Case "241"
 nl="你不在聊天室中，不能进行“级别管理”。"
Case "242"
 nl="你没有执行“升级操作”的权限。"
Case "243"
 nl="你没有执行“降级操作”的权限。"
Case "244"
 nl="用户名不能为空。"
Case "245"
 nl="找不到用户名：<font color=red>" & Request.QueryString("username") & "</font>。"
Case "246"
 nl="该用户名的级别不比你低，不能对其进行操作。"
Case "247"
 nl="选定的级别值不合法。"
Case "248"
 nl="请输入原因。"
Case "250"
 nl="你没有查看“HTML妙用”的权限。"
Case "255"
 nl="你没有更改“站长公告”的权限。"
Case "256"
 nl="内容不能为空。"
Case "260"
 nl="你没有“动作管理”的权限。"
Case "261"
 nl="找不到该动作。"
Case "270"
 nl="你没有“留言管理”的权限。"
Case "271"
 nl="该留言不存在。"
Case "280"
 nl="你没有“调整级别”的权限。"
Case "281"
 nl="你不在聊天室中不能“调整级别”。"
Case "282"
 nl="用户名不能为空。"
Case "283"
 nl="请输入调整级别的原因。"
Case "300"
 nl="你没有管理“投票系统”的权限。"
Case "301"
 nl="不符合投票条件，不能投票。"
Case "302"
 nl="格式错误。"
Case "303"
 nl="请选择你支持的候选人。"
Case "304"
 nl="候选人不存在。"
Case "305"
 nl="候选人不能为空。"
Case "306"
 nl="候选人已经存在，不能重复添加。"
Case "310"
 nl="你没有“永久封锁”ＩＰ的权限。"
Case "311"
 nl="格式错误。"
Case "320"
 nl="IP不能为空。"
Case else
 nl="对不起，该出错类型未被登记。"
End Select
nl="<p>　　" & nl & "</p>"%><html>
<head>
<title>出错提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center"> 
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
              <table border="0" bgcolor="#3399FF" cellspacing="0" cellpadding="2" width="350">
                <tr>
<td width="342"><font color="FFFFFF">¤出错提示</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="关闭"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr> 
                  <td width="59" align="center" valign="top"><font face="Wingdings" color="#0066FF" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value="　确 定　" onclick=javascript:history.go(-1)>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>
