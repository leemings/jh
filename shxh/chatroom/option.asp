<head>
<link rel="stylesheet" href="css1.css">
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
</head>
<body  bgcolor="#FFE4CA" background="../images/bg.gif" oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align=center>
  <table width='100%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF >
    <tbody> 
    <tr> 
      <td align="center" valign="middle"><a href="#" onclick=javascript:window.open('../ballot/ballot.asp','market','left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availheight*3/4+',Status=yes,scrollbars=no,resizable=no'); onmouseover="window.status='请在这儿投票';return true;" onmouseout="window.status='';return true;" title='请在这儿投票'>投票</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="refresh.asp" target="resfrm" onmouseover="window.status='如果你长时间不涮新，请尝试清屏';return true;" onmouseout="window.status='';return true;" title='如果你长时间不涮新，请尝试清屏'>清屏</a></td>
    </tr>
    <tr>
      <td align="center" valign="middle"><a href=# onClick="javascript:window.open('../readme.htm','readme',' width=500,height=300,left=200,top=200,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=新手入门>帮助</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="#" onClick="MM_openBrWindow('announce.htm','gg','width=380,height=400')">公告</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="javascript:parent.listmask();" target="resfrm" onmouseover="window.status='如果你不喜欢某人的发言，可以使用屏蔽功能';return true;" onmouseout="window.status='';return true;" title='如果你不喜欢某人的发言，可以使用屏蔽功能'>屏蔽</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="onlinelist.asp" target="resfrm" onmouseover="window.status='显示在线名单';return true;" onmouseout="window.status='';return true;" title='显示在线名单'>在线</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="image.asp" target='resfrm' onmouseover="window.status='发送图片给你的朋友';return true;" onmouseout="window.status='';return true;" title='发送图片给你的朋友'>贴图</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="websong.asp" target='resfrm' onmouseover="window.status='好歌送好友';return true;" onmouseout="window.status='';return true;" title='好歌送好友'>点歌</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="saveexp.asp" target='resfrm' onmouseover="window.status='储存当前进度';return true;" onmouseout="window.status='';return true;" title='储存当前进度'>储存</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="corp.asp" target="resfrm" onmouseover="window.status='选择你想加入的帮派';return true;" onmouseout="window.status='';return true;" title='选择你想加入的帮派'>加入</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="javascript:parent.talkfrm.settalk('//查看','')" target="talkfrm" onmouseover="window.status='查看他人基本状态';return true;" onmouseout="window.status='';return true;" title='查看他人基本状态'>查看</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="leftcorp.asp" target="resfrm" onmouseover="window.status='离开你现在所在的门派';return true;" onmouseout="window.status='';return true;" title='三思而后行，这可是大罪呀'>叛帮</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="getmoney.asp" target="resfrm" onmouseover="window.status='到你所在的帮派领钱，多少视你的积分和等级还有帮派而定';return true;" onmouseout="window.status='';return true;" title='每天只能领一次哟，这就是你的工资'>领钱</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="seeabout.asp" target="resfrm" onmouseover="window.status='查看自己当前的状态';return true;" onmouseout="window.status='';return true;" title=''>状态</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="seeequip.asp" target="resfrm" onmouseover="window.status='查看自己当前的装备';return true;" onmouseout="window.status='';return true;" title='查看自己的装备'>装备</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="seecurative.asp" target="resfrm" onmouseover="window.status='查看自己拥有的药品';return true;" onmouseout="window.status='';return true;" title='查看自己拥有的药品'>药品</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="miniweapon.asp" target="resfrm" onmouseover="window.status='查看自己拥有的暗器';return true;" onmouseout="window.status='';return true;" title='查看自己拥有的器暗'>暗器</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="attack.asp" target="resfrm" onmouseover="window.status='使用你学过的招式攻击他人';return true;" onmouseout="window.status='';return true;" title='使用你学过的招式攻击他人'>攻击</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="learn.asp" target="resfrm" onmouseover="window.status='提升你业已学会的武功或允许你修习的本门功夫';return true;" onmouseout="window.status='';return true;" title='提升你业已学会的武功或允许你修习的本门功夫'>修习</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="webicq.asp" target="resfrm" onmouseover="window.status='允许呼叫所有在线的朋友';return true;" onmouseout="window.status='';return true;" title='允许呼叫所有在线的朋友'>呼叫</a></td>
    </tr>
    <tr> 
	  <td align="center" valign="middle"><a href="card.asp" target="resfrm" onmouseover="window.status='会员功能开放时，只有会员能使用道具';return true;" onmouseout="window.status='';return true;" title='会员功能开放时，只有会员能使用道具'>道具</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="#" onclick=javascript:window.open('../other/top.asp','top','left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availheight*3/4+',Status=yes,scrollbars=yes,resizable=no'); onmouseover="window.status='风云排行榜';return true;" onmouseout="window.status='';return true;" title='风云排行榜'>排行</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="#" onclick=javascript:window.open('../street/armorshop.asp','armorshop','left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availheight*3/4+',Status=yes,scrollbars=yes,resizable=no'); onmouseover="window.status='武器商店';return true;" onmouseout="window.status='';return true;" title='武器商店'>武具</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="#" onclick=javascript:window.open('../street/curativeshop.asp','curativeshop','left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availheight*3/4+',Status=yes,scrollbars=no,resizable=no'); onmouseover="window.status='在这儿可以购买药品';return true;" onmouseout="window.status='';return true;" title='在这儿可以购买药品'>药店</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="#" onclick=javascript:window.open('../street/prison.asp','prison','left=100,top=50,width='+screen.availWidth*2/3+',Height='+screen.availheight*3/4+',Status=yes,scrollbars=no,resizable=no'); onmouseover="window.status='监牢重地，闲人免进';return true;" onmouseout="window.status='';return true;" title='监牢重地，闲人免进'>监狱</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="#" onclick=javascript:window.open('../admin/onlinelist.asp','manage','Left=100,Top=100,Width='+screen.availwidth*2/3+',Height='+screen.availHeight*3/4+',status=yes,scrollbars=yes,resize=no'); onmouseover="window.status='在线聊务管理';return true;" onmouseout="window.status='';return true;" title='在线聊务管理'>管理</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle"><a href="protect.asp" target="resfrm" onmouseover="window.status='会员功能开放时，只有会员能申请保护';return true;" onmouseout="window.status='';return true;" title='会员功能开放时，只有会员能申请保护'>保护</a></td>
    </tr>
    <tr> 
      <td align="center" valign="middle" bgcolor="#FFFFFF"><a href="#" onclick='javascript:exitchat();' onmouseover="window.status='退出聊天室';return true;" onmouseout="window.status='';return true;" title='退出聊天室'>退出</a></td>
    </tr>
    </tbody> 
  </table>
</div>
</body>

