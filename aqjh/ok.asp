<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("aqjh_chatimage")
chatroombgcolor=Session("afa_chatbgcolor")
Select Case id
Case "100"
nl="成功修改密码，新密码为 <font color=red>" & Request.QueryString("new") & "</font>，请记牢。"
Case "101"
nl="自杀操作完成！用户名 <font color=red>" & Request.QueryString("name") & "</font> 已经被删除，该用户名的所有记录及权限均已消失。"
Case "110"
nl="个人信息已经修改成功。"
Case "111"
nl="你改名操作完成！"
Case "200"
nl="消息已经成功发送给<font color=red>" & Request.QueryString("name") & "</font>。可以在消息列表中查看该消息是否已经被取出。"
Case "210"
nl="动作已经添加完成。"
Case "300"
nl="数据库备份完成，请到 backup 目录下载 global.mdb 进行压缩。"
Case "301"
nl="已经用压缩后的数据库覆盖旧数据库，确认证确后，请删除该备份数据库，以防被他人下载。"
Case "302"
nl="备份数据库删除完成。"
Case "600"
nl="结婚登记成功！收取费用1000两！"
Case "601"
nl="经过漫长的洗浴，你发现自己清爽了不少，精神也好多了，也许该到家里睡一觉会更加开心吧。"
Case "700"
nl="恭喜你！数据库更新成功！"
Case "701"
nl="门派资料已成功修改！"
Case "702"
nl="专用包设置完成！请删除 setup.asp 文件，以免重复运行！"
Case "703"
nl="掌门被开除了！"
Case "704"
nl="删除成功！！"
Case "705"
nl="江湖告示：官符救灾工作顺利完工！"
Case "706"
nl="江湖告示：官符收缴百姓药器的工作顺利完工！"
Case "707"
nl=""
Case else
nl="对不起，该类型未被登记。"
End Select
nl="<p>　　" & nl & "</p>"%><html>
<head>
<title>操作成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false  background=JHIMG/Bk_hc3w.gif leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#dcd8d0">
<tr>
<td>
<table border="0" cellspacing="0" cellpadding="2" width="370" background="jhimg/title.jpg">
<tr>
<td width="344"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">操作成功</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="22">
<tr>
<td width="18"><b><a href="<%if id="200" then%>javascript:window.close();<%else%>javascript:history.go(-1)<%end if%>" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="返回"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="348" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="267"> <%=nl%> </td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<%if id="200" then
Response.Write "<input type='button' name='ok' value=' 查看列表 ' onclick=javascript:top.location.href='chat/webicqlist.asp'>"
else
Response.Write "<input type='button' name='ok' value='　确 定　' onclick='javascript:window.close()'>"
end if%>
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