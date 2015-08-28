<%
username=session("yx8_mhjh_username")%>
<html>
<head>
<script language=javascript>
<!--
function exitchat(){
if(confirm('你确信退出吗？')){
parent.resfrm.location.href='exitchat.asp';
}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
window.open(theURL,winName,features);
}
//-->
</script>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新起点</title>
<style type='text/css'></style>
<link rel="stylesheet" href="css.css" type="text/css">
<base target="_self">
</head>
<body background=bg1.gif oncontextmenu=self.event.returnValue=false leftmargin="0" marginwidth="0" topmargin="2" marginheight="0">
<table width="142" cellpadding="1" cellspacing="1" border="1" align="center" bordercolor="#A75454" class="p9" height="85">
  <tr align="center"> 
    <td bgcolor="#FFFFFF" width="34" height="17"><a href="attack.asp" target="resfrm" onmouseover="window.status='使用你学过的招式攻击他人';return true;" onmouseout="window.status='';return true;" title='使用你学过的招式攻击他人'>攻击</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="miniweapon.asp" target="resfrm" onmouseover="window.status='查看自己拥有的暗器';return true;" onmouseout="window.status='';return true;" title="查看自己拥有的器暗" style="font-size: 9pt; color: black; text-decoration: none">
    暗器</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
        <a href="../t/xw/gj.asp" target="resfrm" onmouseover="window.status='随师傅习武';return true;" onmouseout="window.status='';return true;" title="随师傅习武" style="color: #000000; cursor: e-resize; font-family: 宋体; font-size: 9pt; text-decoration: none">
        随师</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="learn.asp" target="resfrm" onMouseOver="window.status='提升你业已学会的武功或允许你修习的本门功夫';return true;" onMouseOut="window.status='';return true;" title="提升你业已学会的武功或允许你修习的本门功夫">练武</a></td>
  </tr>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" width="34" height="17">
    <a href="seeequip.asp" target="resfrm" onmouseover="window.status='查看自己当前的装备';return true;" onmouseout="window.status='';return true;" title="查看自己的装备" style="font-size: 9pt; color: black; text-decoration: none">
    装备</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="seecurative.asp" target="resfrm" onmouseover="window.status='查看自己拥有的药品';return true;" onmouseout="window.status='';return true;" title="查看自己拥有的药品" style="font-size: 9pt; color: black; text-decoration: none">
    药品</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="card.asp" target="resfrm" onmouseover="window.status='会员功能开放时，只有会员能使用道具';return true;" onmouseout="window.status='';return true;" title="会员功能开放时，只有会员能使用道具" style="font-size: 9pt; color: black; text-decoration: none">
    道具</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="seeabout.asp" target="resfrm" onmouseover="window.status='查看自己当前的状态';return true;" onmouseout="window.status='';return true;" title=''>状态</a></td>
  </tr>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" width="34" height="17"><a href="../t/ys/gj.asp" target="resfrm" onmouseover="window.status='使用轩辕武学攻击他人';return true;" onmouseout="window.status='';return true;" title='使用轩辕武学攻击他人'>攻打</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="yaobu/breed.asp" target="resfrm" onMouseOver="window.status='给你所在的门派补充体力';return true;" onMouseOut="window.status='';return true;" title='给你所在的门派补充体力'>补给</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="corp1.asp" target="resfrm" onMouseOver="window.status='选择你想攻打的要塞';return true;" onMouseOut="window.status='';return true;" title='选择你想攻打的要塞'>要塞</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="PROTECT.ASP" target="resfrm"" onmouseover="window.status='掌门给自己的门派开启保护';return true;" onmouseout="window.status='';return true;" title='掌门给自己的门派开启保护'> 
      门保</a></td>
  </tr>
  <tr align="center"> 
    <td height="17" bgcolor="#FFFFFF" width="34"><a href="onlinelist.asp" target="resfrm" onmouseover="window.status='显示在线名单';return true;" onmouseout="window.status='';return true;" title='显示在线名单'>在线</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="saveexp.asp" target='resfrm' onmouseover="window.status='储存当前进度';return true;" onmouseout="window.status='';return true;" title='储存当前进度'>存点</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="webicq.asp" target="resfrm" onmouseover="window.status='允许呼叫所有在线的朋友';return true;" onmouseout="window.status='';return true;" title="允许呼叫所有在线的朋友" style="font-size: 9pt; color: black; text-decoration: none">
    呼叫</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="image.asp" target="resfrm" onmouseover="window.status='发送图片给你的朋友';return true;" onmouseout="window.status='';return true;" title="发送图片给你的朋友" style="font-size: 9pt; color: black; text-decoration: none">
    贴图</a></td>
  </tr>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" width="34" height="17">
    <a href="st.asp" target="resfrm" onmouseover="window.status='在线看电视';return true;" onmouseout="window.status='';return true;" title="在线看电视" style="font-size: 9pt; color: black; text-decoration: none">
电视</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="song.asp" target="resfrm" onmouseover="window.status='好歌送好友';return true;" onmouseout="window.status='';return true;" title="好歌送好友">
    点歌</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
        <a href="#" onclick="javascript:window.open('zs/index.asp','manage','Left=100,Top=100,Width='+screen.availwidth*3/4+',Height='+screen.availHeight*2/3+',status=no,scrollbars=no,resize=no');" onmouseover="window.status='转生';return true;" onmouseout="window.status='';return true;" title="转生">
        转生</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35">
    <a href="#" onClick="window.open('../garden/hua.htm','reg','width=800,height=550,scrollbars=yes')">
    花园</a></td>
  </tr>
  <tr align="center"> 
    <td height="17" bgcolor="#FFFFFF" width="34">
    <a href="javascript:parent.listmask();" target="resfrm" onmouseover="window.status='如果你不喜欢某人的发言，可以使用屏蔽功能';return true;" onmouseout="window.status='';return true;" title="如果你不喜欢某人的发言，可以使用屏蔽功能" style="font-size: 9pt; color: black; text-decoration: none">
    屏蔽</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="refresh.asp" target="resfrm" onmouseover="window.status='如果你长时间不涮新，请尝试清屏';return true;" onmouseout="window.status='';return true;" title='如果你长时间不涮新，请尝试清屏'>清屏</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="flower.asp" target="resfrm" onmouseover="window.status='选择你想送的花';return true;" onmouseout="window.status='';return true;" title='选择你想想送的花,送给谁'>送花</a></td>
    <td height="17" bgcolor="#FFFFFF" width="35"><a href="#" onclick='javascript:exitchat();' onmouseover="window.status='退出聊天室';return true;" onmouseout="window.status='';return true;" title='退出聊天室'>退出</a></td>
  </tr>
</table>