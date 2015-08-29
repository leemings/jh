<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>关于本站</title>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type='text/css'>
body{font-size:9pt;
scrollbar-face-color:#ffffff; 
scrollbar-shadow-color:#006633; 
scrollbar-highlight-color:#006633;
scrollbar-3dlight-color:#ffffff;
scrollbar-darkshadow-color:#ffffff;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#006633;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body bgcolor="#003366" leftmargin="0" topmargin="0">
<table border="0" width="520" style="text-decoration: blink; color: #008080; line-height: 150%; border: 1 solid #009933" bgcolor="#FFFFFF" align="center" cellpadding="5" cellspacing="5">
  <tr> 
    <td width="100%" align="center" bgcolor="#006633"> <div align="center"><b><font color="#ffff00"><font color="#ffff00">关 
        于 本 站</font></font></b></div>
      
    </td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#003366" > <font color="#008080"><b><font color="#FF0000">-站点说明</font></b></font><font color="#FF0000">-</font><br>
      <table width="100%" border="0" cellspacing="2" cellpadding="3">
	  <tr> 
          <td width="15%"><div align="center"><font color="#00FF00">江湖名称</font></div></td>
          <td width="85%"><font color="#FFFFFF">『E线江湖』</font></td>
        </tr>
		<tr> 
          <td width="15%"><div align="center"><font color="#00FF00">江湖地址</font></div></td>
          <td width="85%"><a href="http://51eline.com" target="_blank"><font color="#FFFFFF">http://51eline.com</font></a></td>
        </tr>
		<tr> 
          <td width="15%"><div align="center"><font color="#00FF00">论坛地址</font></div></td>
          <td width="85%"><a href="http://bbs.51eline.com" target="_blank"><font color="#FFFFFF">http://bbs.51eline.com</font></a></td>
        </tr>
        <tr> 
          <td width="15%"><div align="center"><font color="#00FF00">江湖站长</font></div></td>
          <td width="85%"><font color="#FFFFFF">一线天</font></td>
        </tr>
        <tr> 
          <td width="15%"><div align="center"><font color="#00FF00">江湖程序</font></div></td>
          <td width="85%"><font color="#FFFFFF">由一线天独立改版美化.使用权归一线天个人所有.</font></td>
        </tr>
        <tr> 
          <td width="15%"><div align="center"><font color="#00FF00">开通日期</font></div></td>
          <td width="85%"><font color="#FFFFFF">2003.3.16</font></td>
        </tr>
        <tr> 
          <td width="15%" valign="top">
<div align="center"><font color="#00FF00">关于会员</font></div></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">所有注册用户初始化一级会员.银两1000万.泡点是正常的1.2倍.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">三个月内若表现良好、等级达60级可继续升为二级会员.泡点是正常的1.4倍.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">之后达90级可升为三级会员.泡点是正常的1.6倍.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">达120级可升为四级会员.泡点是正常的1.8倍.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">180级可升为泡点制会员.泡点是正常的2倍.</font></td>
              </tr>
			                <tr>
                <td><font color="#FFFFFF">达300级可升为五级会员.泡点是正常的2.5倍.</font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="15%" valign="top">
<div align="center"><font color="#00FF00">江湖特色</font></div></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">来者皆会员</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">有4个登陆页面.优美的背景音乐衬托主题.每次登陆随机出现.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">登陆后的页面仿亲切的Windows视窗.16种图标、35首背景音乐随机出现.</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">有20种鼠标样式以及各种漂亮的滚动条.你的浏览器IE版本是6.0的吗.</font></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#003366" > <font size="2"><b><font color="#FF0000">-更新记录-</font></b></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="15%"><font color="#00CC00">2003-03-16</font></td>
          <td width="85%"><font color="#FFFFFF">开通测试</font></td>
        </tr>
        <tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-04</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><div align="center"><font color="#FF6600"><strong>第一次升级</strong></font></div></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正拜师习武修炼过快</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正掉鱼无法使用</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正掌门发放功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正状态无法显示具体上限和目前状态数值</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正倒夜香错误</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正扒点问题</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正会员工资领取数量(以前领的太多了，花都花不完）</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增粪库、法力排行等等</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加解除睡眠</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖医院</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加旧车交易</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加爱情小屋</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增强系统安全性</font></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td width="15%"><font color="#00CC00">2003-04-27</font></td>
          <td width="85%"><font color="#FFFFFF">改成桌面版、图标16种随机</font></td>
        </tr>
        <tr>
          <td width="15%"><font color="#00CC00">2003-05-01</font></td>
          <td width="85%"><font color="#FFFFFF">增加江湖绿色QQ（正常浏览需IE6.0）</font></td>
        </tr>
        <tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-03</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><div align="center"><font color="#FF6600"><strong>第二次升级</strong></font></div></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加打扑克、打麻将功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖神秘花园</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖行走功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加图形装备界面</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加领养小孩、喂养小孩、生育小孩</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加孤岛冒险</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖宝物信息</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加飞花舞剑</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加在线大富翁</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加游戏中心</font></td>
              </tr>
			   <tr>
                <td><font color="#FFFFFF">美化部分功能页面</font></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td width="15%"><font color="#00CC00">2003-05-05</font></td>
          <td width="85%"><font color="#FFFFFF">修正当铺功能、增加武林大会</font></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-12</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><div align="center"><font color="#FF6600"><strong>第三次升级</strong></font></div></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖怨偶</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖飞行</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖航海</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖拜神</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加新版江湖股票</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖淘金</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖青楼</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖鸭店</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加江湖逃婚</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加属性转换</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加我的徒弟、师傅发钱、师傅指导</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加门派议事</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加掌门让位</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加太白酒楼</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加抗击倭寇</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加抢劫银行</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加转换市场</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加法力转换</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加轻功转换</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加领取工资新方法</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加密码保护</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加金木水火土属性</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加PK功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加银币</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加发放狗狗、舞女、舞男、武士、八国盟军、抢存点、豆点、金卡等</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加职业[小偷] 偷金币</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加寻金币、寻银币</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加招收官府</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加原子弹</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加二号情人、三号情人</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加掌门超级令</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">后台增加武功设置(可以调改任何门派的武功字眼)</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">江湖注册增加认证码</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正孤岛冒险无时间限制</font></td>
              </tr>
			   <tr>
                <td><font color="#FFFFFF">江湖排行新增金木水火土属性、充值卡、银币、帅哥、靓女排行</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">... ...</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-13</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">增加表情库</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加歌库</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">增加宠物自爆</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正、加强点歌系统</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-14</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">增加江湖论坛</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加自我保释</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">增加睡眠自救</font></td>
              </tr>
            </table></td>
        </tr>
		
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-15</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">点歌系统增加19首经典歌曲</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">表情库增加70个小动画</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-16</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">新增会员修炼</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增会员学武</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">新增会员钓鱼</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增会员挖宝</font></td>
              </tr>
			  <tr>
                <td><font color="#FFFFFF">为江湖各类事件添加动感声效</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%"><font color="#00CC00">2003-05-21</font></td>
          <td width="85%"><font color="#FFFFFF">聊天室界面全面改版，塑造『E线江湖』特有风格</font></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-24</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">添加聊天室顶部的问候语</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">聊天室发言区的功能作了一些调整</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%"><font color="#00CC00">2003-05-26</font></td>
          <td width="85%"><font color="#FFFFFF">新增站长、系统发放怪鸟、饿鼠、猪哥、卡片、配药、苹果、别墅、礼物</font></td>
        </tr>
		<tr>
          <td width="15%">&nbsp;</td>
          <td width="85%"><font color="#FFFFFF">新增系统随机任务：捉人贩、抓小偷</font></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-05-27</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">增加卡片:免踢卡,万能卡,情人卡,分手卡</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加会员绝招</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%"><font color="#00CC00">2003-05-28</font></td>
          <td width="85%"><font color="#FFFFFF">新增会员等级自动转换：60级升2级会员、90级升3级会员、120级升4级会员</font></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-05</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">新增申冤功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增轩辕（个人武功）</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-06</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修改进入聊天室的公告</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加垂直分屏</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">增加在线广播</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-09</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">论坛与江湖整合</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">加强种花功能</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">增加收花功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加系统发放金币、存点、豆点</font></td>
              </tr>
			                <tr>
                <td><font color="#FFFFFF">增加系统出题数两百多道</font></td>
              </tr>
            </table></td>
        </tr>
	<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-11</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">新增夺宝大赛</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增门派论剑</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">新增魔宝功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增屏蔽功能</font></td>
              </tr>
			                <tr>
                <td><font color="#FFFFFF">新增对赌银两</font></td>
              </tr>
			  <tr>
                <td><font color="#FFFFFF">新增邀请舞伴</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">加强邀请跳舞功能</font></td>
              </tr>
			                <tr>
                <td><font color="#FFFFFF">增加发射子弹</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-16</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正修炼紫金葫芦所出现的问题</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">增加5级会员</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">增加系统自动提示：PK（比武）开始</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-24</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">新增在线视听功能</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">加强门派大战功能</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-06-28</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">江湖论坛升级（经过两次）为：最新版 带最新QQ个人秀</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-03</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">江湖个人状态显示论坛里设计的个人秀</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-05</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">江湖聊天室名单区状态显示更改、新增右键显示个人秀</font></td>
              </tr>
            </table></td>
        </tr>
			<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-14</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">论坛新增形象大赛、结婚礼堂、多人台球功能</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-18</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">论坛添加头像集锦和美丽图库（数十万张）
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-22</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">应不少朋友的要求，E线江湖7.8.0提供下载，包含现有江湖的大部分功能

</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-25</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">美化聊天室里的超酷音乐系统
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-26</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">论坛新增多功能日记本、在线收藏夹、智力游戏
</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">聊天室的超酷音乐系统增加133首歌曲（其中有百首韩曲）</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-07-31</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正自创门派不能用的问题
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-01</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正不看详细状态就没有攻击和防御上限的问题
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-06</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正扑克、麻将不能用的问题
</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修正超级令不能用的问题</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-07</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正门派聚议弟子申请长老后管理等级问题
</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">修改掌门招收长老、护法、堂主数量有误的问题</font></td>
              </tr>
<tr>
                <td><font color="#FFFFFF">修正长老不能招收弟子的问题</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-09</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">修正单身女子可以生小孩问题
</font></td>
              </tr>
              <tr>
                <td><font color="#FFFFFF">新增周公解梦功能</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-13</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">完成{一线网络}主页的制作
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-15</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作和发布原创插件：一线视听--音乐听吧
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-22</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">完成E线江湖主页第三版（清新靓丽）的制作
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-08-27</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">新增DJ舞曲947首
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-09-15</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">点歌系统曲目新增2万多首
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-09-18</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作并发布免虚文件
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-09-19</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作并发布原创插件[Flash频道]
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-09-23</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作并发布原创插件[DJ舞吧]
</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-09-27</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">重新美化原创插件[DJ舞吧]

</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-10-04</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作并发布[蓝色嘭嘭]风格和选择风格式登陆口

</font></td>
              </tr>
            </table></td>
        </tr>
		<tr>
          <td width="15%" valign="top"><font color="#00CC00">2003-10-05-07</font></td>
          <td width="85%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td><font color="#FFFFFF">制作并发布[童话世界、骑士在线、猎人传说、一线使命、圣者无敌、怪兽乐园、宠物之王]等7种风格版式

</font></td>
              </tr>
            </table></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#003366"> 
      <p><font color="#FF0000" size="2"><b>-江湖规矩-</b></font><font size="2"><br>
        </font><font color="#FFFFFF">1.江湖中不允许骂人.说脏话.杀人不犯法.骂人要坐牢.<br>
        2.不要粘贴大量重复的话一次性发出.也不允许多次发送一句重复的话.<br>
        3.不允许在江湖中拉人</font> <font color="#FFFFFF">...... 最好不要在江湖中谈论其他的江湖(特别是站长禁止的网址)<br>
        <font color="#FFFF00">上面这些基本上是会立即被系统自动踢出的.</font></font><br>
        <font color="#FFFFFF">4.其他的最好是不要威胁恐吓站长.在聊天室中别人至少有几百种方法让你倒霉.<br>
        5.不要试图做不该做的事.访问系统管理页面或进行破坏.『E线江湖』的这些地方都设置了炸弹保护.悠着点.</font> <br>
        <font color="#FFFFFF">6.做官府的管理员一定要尽职尽责.公正无私.每个管理员的操作都会被记录在管理日志中.供大家临督.<br>
        7.官府人员只管理刷屏、骂人.及违反中国有关互联网规定的对象.可以使用警告、逮捕、封锁IP的手段制止恶意捣乱的行为的发生.</font> <font color="#FFFFFF"><br>
        8.官府人员不得参与江湖恩怨.不可以偏向任何一方.只要没有捣乱行为出现.杀人的事情一般不应过问——滥杀新手除外.欺负新朋友的人一定要加以惩罚！</font> 
        <font color="#FFFFFF"><br>
        9.如果官府人员很长时间不能上网.影响了自己的工作和秩序.我们会考虑选举新的管理人员来顶替.<font color="#00FF00">原则上要求每天在聊天室的时间不少于3个小时</font>.</font></p></td>
  </tr>
</table>
</body>
</html>
