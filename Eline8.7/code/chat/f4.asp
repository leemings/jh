<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>♀一线网络→wWw.51eline.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=javascript>
minimizebar='ico';
closebar="close.gif"; 
icon="icon.gif"; 
function noBorderWin(fileName,w,h,titleBg,titleColor,titleWord,borderColor,scr)
{
  nbw=window.open('','','fullscreen=yes');
  nbw.resizeTo(w,h);
  nbw.moveTo((screen.width-w)/2,(screen.height-h)/2);
  nbw.document.writeln("<title>"+titleWord+"</title>");
  nbw.document.writeln("<"+"script language=javascript"+">"+"cv=0"+"<"+"/script"+">");
  nbw.document.writeln("<body topmargin=0 leftmargin=0 scroll=no>");
  nbw.document.writeln("<table cellpadding=0 cellspacing=0 width=100% height=100% style=\"border: 1px solid "+borderColor+"\">");
  nbw.document.writeln("  <tr bgcolor="+titleBg+" height=20 onselectstart='return false' onmousedown='cv=1;x=event.x;y=event.y;setCapture();' onmouseup='cv=0;releaseCapture();' onmousemove='if(cv)window.moveTo(screenLeft+event.x-x,screenTop+event.y-y);'>");
  nbw.document.writeln("    <td style=\"font-family:宋体; font-size:12px; color:"+titleColor+";\"><span style='top:1;position:relative;cursor:default;'>"+titleWord+"</span></td>");
  nbw.document.writeln("    <td align=right width=30 style=font-size:16px;>");
  nbw.document.writeln("      <img border=0 width=12 height=12 alt=最小化 src="+minimizebar+" onclick=window.blur();>");
  nbw.document.writeln("      <img border=0 width=12 height=12 alt=关闭 src="+closebar+" onclick=window.close();>");
  nbw.document.writeln("    </td>");
  nbw.document.writeln("  </tr>");
  nbw.document.writeln("  <tr>");
  nbw.document.writeln("    <td colspan=3><iframe frameborder=0 width=100% height=100% scrolling="+scr+" src="+fileName+"></iframe></td>");
  nbw.document.writeln("  </tr>");
  nbw.document.writeln("</table>");
}
</SCRIPT>

<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;CURSOR:url('1.cur');}
</style>
<script Language="Javascript">
function s(){parent.f2.document.af.mdsx.checked=false;parent.f2.document.af.sytemp.focus();}
if(parent.document.URL.indexOf("file:")>=0||parent.f2.document.URL.indexOf("file:")>=0){parent.location.href='chaterr.asp?id=001';}
function showname(){
parent.f2.document.af.mdsx.checked=true;parent.f2.document.af.sytemp.focus();
if(parent.m.location.href=="about:blank"){
parent.m.location.href="f3.asp";}
else
{parent.m.location.reload();}
}
function ex(){parent.t.location.href='about:blank';top.location.href="exitlt.asp";return true;}
</script>

</head>

<body background="f2.gif" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td valign="top" background="]">
	<table width="160" height="20" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#E1D7A2" bordercolordark="#efefef" background="yinyue/bg2.gif" bgcolor="#efefef">
        <tr>
          <td valign="top"> 
            <div align="center"> 
              <IFRAME src="yinyue/bgm.htm" width="150" height="15" scrolling="no" frameborder="0"></iframe>
            </div></td>
  </tr>
</table>
	  <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="f2.gif">
        <tr align="center"> 
          <td> <a href="selectchatroom.asp" onClick="javascript:s()" onMouseOver="window.status='更换房间';return true" onMouseOut="window.status='';return true" target="f3" title="选择房间">房间</font></a></td>
          <td><a href="#" onMouseOver="window.status='立即更新在线聊友名单。';return true" onMouseOut="window.status='';return true" onClick="javascript:showname();" title="刷新在线聊友名单">名单</a></td>
          <td><a href="../CHANGAN/xiaowu.asp" onClick="javascript:s()" target="_blank" onMouseOver="window.status='爱情小屋';return true" onMouseOut="window.status='';return true" title="爱情小屋">小屋</a></td>
          <td><a href="npc/npc.asp" onclick="javascript:s()" title="行走功能" target="f3">行走</a></td>
        </tr>
        <tr align="center"> 
          <td> <a href="caidan.asp" onClick="javascript:s()" onMouseOver="window.status='功能库';return true" onMouseOut="window.status='';return true" target="f3" title="功能库"><font color="#00FFFF">菜单</font></a></td>
          <td><a href="KiNg.asp" onClick="javascript:s()" onMouseOver="window.status='更多命令使用';return true" onMouseOut="window.status='';return true" target="f3" title="江湖设施与常用功能">功能</font></A></td>
          <td><a href='command.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='动作。';return true" onMouseOut="window.status='';return true" title="发一下动作">动作</font></a> 
          </td>
          <td> <a href="cw/cw.asp" onClick="javascript:s()" target=f3 title="照顾自己的小动物">宠物</font></a></td>
        </tr>
        <tr align="center"> 
          <td><a href="savevalue.asp" onClick="javascript:s()" onMouseOver="window.status='立即保存经验值，并显示当前状态。';return true" onMouseOut="window.status='';return true" target="f3" title="查看当前存点">存点</font></a></td>
          <td><a href="game/game_index.htm" onclick="javascript:s()" title="精心挑选的数十款好玩小游戏" target="f3">游戏</font></a> 
            <a href='setfontsize.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='改变对话区字体的磅值及行距，设置完成后必须点“清屏”才能起作用。';return true" onMouseOut="window.status='';return true" title="改变字体大小及行距"></a></font></td>
          <td><a href="BOY/Boy.asp" onClick="javascript:s()" target="f3" title="细心照顾您的小孩">小孩</font></a> 
          </td>
          <td> <a href="#" onClick="window.open('../hcjs/jhjs/vehicle.asp','','scrollbars=yes,resizable=yes,width=670,height=400')" title="购买交通工具">交通</font></a></td>
        </tr>
        <tr align="center"> 
          <td height="20"><a href="wupin.asp" onClick="javascript:s()" onMouseOver="window.status='看看你现在有哪些物品';return true" onMouseOut="window.status='';return true" target="f3" title="看看你有什么物品">物品</font></a></td>
          <td height="20"><a href='setwg3.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='发招。';return true" onMouseOut="window.status='';return true" title="你派的武功">动武</font></a></td>
          <td height="20"><a href='pic.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='改变对话区字体的磅值及行距，设置完成后必须点“清屏”才能起作用。';return true" onMouseOut="window.status='';return true" title="表情库">图库</a></td>
          <td height="20"><a href="xlwupin.asp" onClick="javascript:s()" onMouseOver="window.status='拍卖物品。';return true" onMouseOut="window.status='';return true" target="f3" title="在线拍卖">拍卖</font></a></td>
        </tr>
        <tr align="center"> 
          <td><a href="#" onClick="javascript:parent.qp()" onMouseOver="window.status='清除对话区的所有对话。';return true" onMouseOut="window.status='';return true" title="请及时清除聊天记录">清屏</font></a></td>
          <td><a href='killman.asp' onClick="javascript:s()" onMouseOver="window.status='查看江湖坏人。';return true" onMouseOut="window.status='';return true" target="f3")" title="查看江湖坏人">恶人</font></a></td>
          <td><a href="song.asp" onMouseOver="window.status='点歌给自己或其他聊友。';return true" onMouseOut="window.status='';return true" target="f3" title="经典歌曲">歌库</font></a></td>
          <td><a href='../10top.asp' onClick="javascript:s()" onMouseOver="window.status='查看江湖10大';return true" onMouseOut="window.status='';return true" target="f3")" title="查看江湖爬行榜">排行</font></a></td>
        </tr>
		<tr align="center"> 
          <td><a href="../tq/Weather.asp" title="全国天气预报" target="f3" onClick="javascript:s()" onMouseOver="window.status='全国天气预报';return true" onMouseOut="window.status='';return true">天气</font></a></td>
          <td><a href='../garden/hua.asp' onClick="javascript:s()" onMouseOver="window.status='看看花养的如何了';return true" onMouseOut="window.status='';return true" target="_blank")" title="看看花养的如何了">花园</font></a></td>
          <td><a href="#" onClick="window.open('../drtq/index.asp','','scrollbars=yes,resizable=yes,width=600,height=400')" title="个人台球!">台球</a></td>
          <td><a href='st.asp' onClick="javascript:s()" onMouseOver="window.status='在线广播和电视';return true" onMouseOut="window.status='';return true" target="f3")" title="音乐不断，给您精彩！">视听</a></td>
        </tr>
		<tr align="center"> 
          <td><a href="../bbs/index.asp" title="来看看论坛，站长不在江湖就肯定在那了" target="_blank" onClick="javascript:s()" onMouseOver="window.status='来看看论坛，站长不在江湖就肯定在那了';return true" onMouseOut="window.status='';return true">论坛</a></td>
          <td><a href='setwg4.asp' onClick="javascript:s()" onMouseOver="window.status='个人轩辕武功';return true" onMouseOut="window.status='';return true" target="f3")" title="个人轩辕武功">轩辕</font></a></td>
          <td><a href='empdz/mp.asp' onClick="javascript:s()" target=f3 onMouseOver="window.status='门派大战';return true" onMouseOut="window.status='';return true" title="门派大战">论剑</font></a></td>
          <td><a href='SETWGdb.ASP' onClick="javascript:s()" onMouseOver="window.status='门派大战，争夺江湖至宝';return true" onMouseOut="window.status='';return true" target="f3")" title="夺宝大赛时专用的武功">夺宝</a></td>
        </tr>
        <tr align="center"> 
          <td><a href="../zhuangtai.asp" onClick="javascript:s()" target="f3" title="我的状态"><font color="#FFCCFF">状</font></a><a href="#" title="查看自己的详细状态" onMouseOver="window.status='查看自己的详细状态。';return true" onMouseOut="window.status='';return true" onclick="javascript:window.open('zt/index.asp','zhuantai','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=500,height=530')" target="_self"><font color="#99CCFF">态</font></a> 
          </td>
          <td><a href="stuntlist.asp" onClick="javascript:s()" target="f3" title="使用特技">特技</a></td>
          <td><a href="face.asp" onClick="javascript:s()" target="f3" title="选择你喜爱的头像">头像</a></td>
          <td><a href="#" onMouseOver="window.status='退出聊天室，并自动保存经验值。';return true" onMouseOut="window.status='';return true" onclick="javascript:ex()" title="安全离开并存点">离开</font></a></td>
        </tr>
		</table>   
</td>    
  </tr>      
</table>   
</body>   
</html>   
   
   
   
   
   
   
  
  
  
  
  
  
