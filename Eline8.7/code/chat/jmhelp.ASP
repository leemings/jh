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
<title>聊天界面帮助 - happyjh.com</title>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type='text/css'>
body{font-size:9pt;
scrollbar-face-color:#ffffff; 
scrollbar-shadow-color:#ff9900; 
scrollbar-highlight-color:#ff9900;
scrollbar-3dlight-color:#ffffff;
scrollbar-darkshadow-color:#ffffff;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#ff9900;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body>
<table border="0" width="80%" style="text-decoration: blink; color: #008080; line-height: 150%; border: 1 solid #000000" bgcolor="#FFFFFF" align="center" cellpadding="5" cellspacing="5">
  <tr> 
    <td width="100%" align="center" bgcolor="#FF9900"> <b><font color="#ffff00">1.<a href="jmhelp.asp"><font color="#ffff00">聊天室界面帮助</font></a> 
      2.</font><a href="xinshou.asp"><font color="#FFFF00">江湖入门帮助</font></a> <font color="#ffff00">3.</font><a href="jhgj.asp"><font color="#FFFF00">江湖规矩 
      </font></a><font color="#ffff00">4.</font><a href="canshu.asp"><font color="#FFFF00">江湖设定档案</font></a></b> 
    </td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#FFCC00" > <font color="#FF0000"><b>界面帮助</b></font><br>
      <hr noshade size="1" color=009900 align="left">
      <table width="371" border="1" cellspacing="2" cellpadding="2" bordercolor="#0000FF" bgcolor="#00CC99">
<tr> 
<td width="307" rowspan="3"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20%">
<tr bgcolor="#ddddff" valign="top"> 
<td height="12" bgcolor="#FFFFFF"> 
<div align="center"><font color="#009933">公告区</font></div>
</td>
</tr>
<tr bgcolor="#ddddff" valign="top"> 
<td bgcolor="#eeeeff"> 
<div align="left"><font color="#ff0000">【公共区】</font></div>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="103%">
<tr bgcolor="#eeeeff" valign="top"> 
<td bgcolor="#eeeeff"> <font color="#000000"><br>
…………………………………………<br>
………………………………<br>
……………………………………<br>
…………………………………………</font></td>
</tr>
<tr bgcolor="#eeeeff" valign="top"> 
<td> 
<hr color="#0000CC">
<font color="#ff0000">【房间区】</font><br>
<font color="#009933">【个人区】</font><br>
…………………………………………<br>
………………………………<br>
…………………………
</td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" bordercolordark="#ffffff" bordercolorlight="#000000" height="30%">
<tr bgcolor="#00cc99"> 
<td bgcolor="#00cc99"> 
<div align="center"><font color="#FF0000">功能区</font><font color="#000000"><br>
<br>
</font></div>
</td>
</tr>
</table>
</td>
<td width="58" rowspan="3"> 

<table width="101%" border="0" cellspacing="0" cellpadding="0" bordercolordark="#ffffff" bordercolorlight="#000000" height="170">
<tr align="middle" bgcolor="#00cc99"> 
<td width="53" bgcolor="#00cc99"> 
<p align="center"><font color="#FF0000">名单区</font><font color="#000000"><br>
                    <br>
                    <br>
                    </font></p>
</td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" bordercolordark="#ffffff" bordercolorlight="#000000">
<tr bgcolor="#00cc99"> 
<td> 
<div align="center"><font color="#FF0000">控制区</font><font color="#000000"><br>
<br>
</font></div>
</td>
</tr>
</table>
</td>
</tr>
</table>
      <p>&nbsp;本聊天室界面如下<font color="#000000"><br>
        <font size="2" color="#0000FF">房间区：</font></font><font size="2">选择要去的房间。</font><font color="#000000" size="2"><br>
        <font color="#0000FF">公告区：</font></font><font size="2">这里显示站长公告。<br>
        </font><font color="#0000FF" size="2">对话区：</font><font size="2">用户对话显示区。<br>
        　　　　鼠标在每句话开头的用户名点击，来选择说话对象。<br>
        　　　　显示时：当这句话是对你说的时候，你的名字会自<br>
        　　　　动变成红色，并加〖〗，以提醒你有人在对你说话。</font></p>
<p><font size="2">　　　　分屏显示时对话区分为公共对话区与个人对话区。<br>
　　　　进入聊天室默认的都是分屏显示，上下分屏时下面是个人区，左右分屏时右面是个人区<br>
　　　　<font color="#FF0000">公共对话区：</font>与你相关的发言显示在公共对话区<br>
　　　　<font color="#FF0000">个人对话区：</font>你说的或你说的或者跟你相关的所有信息显示在个人对话区<br>
<br>
        </font><font color="#0000FF" size="2">名单区：</font><font color="#000000" size="2"></font><font size="2">在线用户列表。<br>
　　　　单击可以选定说话对象<br>
　　　　双击可以查看资料</font></p>
      <p><font color="#0000FF">控制区：</font>聊天室功能按键.....</p>
<p><font color="#0000FF" size="2">功能区：</font><font size="2">用户从此处进行发言。并使用各种聊天室功能。下面对本区的部份功能进行说明。
</font>
</td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#FFCC00" ><font size="2"><b><font color="#FF0000">简要说明</font></b></font><br>
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0" background="ico/wfy_i_c_g_b25.gif">
        <tr> 
                      
          <td height="29" background="ico/wfy_i_c_g_b24.gif"> 
            <div align="center"> 
                     <select name='zt' style='font-size:12px; background-color:#E4CCAC'>
<option selected>离席
<option style=background-color:"#FFEEFF" value="/离席$ 正常">正常
<option style=background-color:"#FFEEFF" value="/离席$ 睡觉">睡觉
<option style=background-color:"#FFEEEE" value="/离席$ 吃饭">吃饭
<option style=background-color:"#EEFFEE" value="/离席$ 看书">看书
<option style=background-color:"#FFFFEE" value="/离席$ 挂网">挂网
<option style=background-color:"#EEEEFF" value="/离席$ 想事">想事
<option style=background-color:"#EEFFFF" value="/离席$ 逛街">逛街
<option style=background-color:"#EEFFFF" value="/离席$ 约会">约会
<option style=background-color:"#EEFFFF" value="/离席$ 沉思">沉思
<option style=background-color:"#EEFFFF" value="/离席$ 洗澡">洗澡
</select>  
        
我对                                                                             
<input type="text" name="towho" readonly size="8" value="大家" style="height: 18; text-align: center; background-image: url('ico/wfy_i_c_g_b25.gif'); border-style: groove; border-width: 1">  
<font class=yy3>说</font>                                                                             
<input type=text name='sytemp'  size=35 maxlength=250 style="height: 18; background-image: url('ico/wfy_i_c_g_b25.gif'); border-style: groove; border-width: 1">
              <a href="#"><img src="ico/wfy_i_c_g_b32.gif" width="20" height="14" border="0" align="absmiddle"></a><a href="#"><img src="ico/wfy_i_c_g_b31.gif" width="20" height="14" border="0" align="absmiddle"></a> 
              <input type=submit name='subsay' value='发送' style="height: 18; background-color:E4CCAC;color:000000;border: 1 double" onmouseover="this.style.color='FF0000'" onmouseout="this.style.color='000000'">
            </div></td> 
                </tr> 
              </table><br>
<select name='zt' style='font-size:12px; background-color:#E4CCAC'>
<option selected>离席
<option style=background-color:"#FFEEFF" value="/离席$ 正常">正常
<option style=background-color:"#FFEEFF" value="/离席$ 睡觉">睡觉
<option style=background-color:"#FFEEEE" value="/离席$ 吃饭">吃饭
<option style=background-color:"#EEFFEE" value="/离席$ 看书">看书
<option style=background-color:"#FFFFEE" value="/离席$ 挂网">挂网
<option style=background-color:"#EEEEFF" value="/离席$ 想事">想事
<option style=background-color:"#EEFFFF" value="/离席$ 逛街">逛街
<option style=background-color:"#EEFFFF" value="/离席$ 约会">约会
<option style=background-color:"#EEFFFF" value="/离席$ 沉思">沉思
<option style=background-color:"#EEFFFF" value="/离席$ 洗澡">洗澡
</select>
      就是设置你当时的状态<br><br>
	  <input type="text" name="towho" readonly size="8" value="大家" style="height: 18; text-align: center; background-image: url('ico/wfy_i_c_g_b25.gif'); border-style: groove; border-width: 1">
      就是你要聊天的对象，可以在名单区左键选择或在对话区左键选择。<br><br>
	  <input type=text name='sytemp'  size=35 maxlength=250 style="height: 18; background-image: url('ico/wfy_i_c_g_b25.gif'); border-style: groove; border-width: 1">
      在这里输入你要说的话或所要执行的动作，然后按 
      <input type=submit name='subsay' value='发送' style="height: 18; background-color:E4CCAC;color:000000;border: 1 double" onmouseover="this.style.color='FF0000'" onmouseout="this.style.color='000000'">
      便OK ^_^ 
      <hr noshade size="1" color=009900 align="left">
	  <select name='addwordcolor'  style='font-size:12px; background-color:#E4CCAC'> 
<option style="color:8800FF" value="8800FF" selected>名</option> 
<option style="color:000000" value="000000">名</option> 
<option style="color:0088FF" value="0088FF">名</option> 
<option style="color:0000FF" value="0000FF">名</option> 
<option style="color:000088" value="000088">名</option> 
<option style="color:888800" value="888800">名</option> 
<option style="color:008888" value="008888">名</option> 
<option style="color:008800" value="008800">名</option> 
<option style="color:8888FF" value="8888FF">名</option> 
<option style="color:AA00CC" value="AA00CC">名</option> 
<option style="color:8800FF" value="8800FF">名</option> 
<option style="color:888888" value="888888">名</option> 
<option style="color:CCAA00" value="CCAA00">名</option> 
<option style="color:FF8800" value="FF8800">名</option> 
<option style="color:CC3366" value="CC3366">名</option> 
<option style="color:FF00FF" value="FF00FF">名</option> 
<option style="color:3366CC" value="3366CC">名</option> 
</select>
      就是选择你出现在对话区的名字的颜色。<br><br>
	  <select name='saycolor'  style='font-size:12px; background-color:#E4CCAC'>                                                                          
<option style="color:008888" value="008888" selected>话</option> 
<option style="color:000000" value="000000">话</option> 
<option style="color:0088FF" value="0088FF">话</option> 
<option style="color:0000FF" value="0000FF">话</option> 
<option style="color:000088" value="000088">话</option> 
<option style="color:888800" value="888800">话</option> 
<option style="color:008888" value="008888">话</option> 
<option style="color:008800" value="008800">话</option> 
<option style="color:8888FF" value="8888FF">话</option> 
<option style="color:AA00CC" value="AA00CC">话</option> 
<option style="color:8800FF" value="8800FF">话</option> 
<option style="color:888888" value="888888">话</option> 
<option style="color:CCAA00" value="CCAA00">话</option> 
<option style="color:FF8800" value="FF8800">话</option> 
<option style="color:CC3366" value="CC3366">话</option> 
<option style="color:FF00FF" value="FF00FF">话</option> 
<option style="color:3366CC" value="3366CC">话</option> 
</select>
      就是选择你的话出现在对话区时的颜色。 <br>
	  <hr noshade size="1" color=009900 align="left">
	  <select name='command'" style='font-size:12px; background-color:#E4CCAC'>                                                                          

<option value="" selected>主要功能
<option style="color:#FF6633" value="">==各人==

<option style="color:#FF6633" value="/暂离$ 您好，我不在请等我回来…… ">暂离三国
<option style="color:#FF6633" value="/回来$ ">回到三国 
<option style="color:#FF6633" value="/暴豆$ ">暴威力豆
<option style="color:#FF6633" value="/好友$ 加入">结义兄弟
<option style="color:#FF6633" value="/好友$ 查看">查看兄弟
<option style="color:#FF6633" value="/好友$ 删除">割袍段义
<option style="color:#FF6633" value="/清除$ 要清除的好友姓名">清除兄弟
<option style="color:#FF6633" value="/信息$">查看信息
<option style="color:#FF6633" value="/标题$ 大家好，请多多关照!">发 
标 题 
<option style="color:#FF6633" value="/修练$ ">宝物修练
<option style="color:#FF6633" value="/抢夺宝物$">抢夺宝物
<option style="color:#B000D0" value="">==银行==
<option style="color:#B000D0" value="/存钱$ 0">银行存钱
<option style="color:#B000D0" value="/取钱$ 0">银行取钱
<option style="color:#B000D0" value="/转账$ 10000">银行转账
<option style="color:#FF0000" value="">==赠送==
<option style="color:#FF0000" value="/丢弃$ 100000">丢弃银两
<option style="color:#FF0000" value="/送钱$ 1000">赠送现金
<option style="color:#FF0000" value="/送豆点$ 10">赠送豆点
<option style="color:#FF0000" value="/送金币$ 10">赠送金币
<option style="color:#FF0000" value="/赠送$ 类别|名字|数量">赠送物品
<option style="color:#FF0000" value="/用卡$ 卡片名">使用卡片
<option style="color:#0808CF" value="">==门派==
<option style="color:#0808CF" value="/加入门派$">加入门派
<option style="color:#0808CF" value="/本门发钱$">本门发钱
<option style="color:#0808CF" value="/指导弟子$">指导弟子
<option style="color:#0808CF" value="/基金$">本门基金
<option style="color:#0808CF" value="/收徒$ ">招收徒弟
<option style="color:#0808CF" value="/逐出$ 徒弟名">逐出师门
<option style="color:#0808CF" value="/拜师$ ">拜师习武
<option style="color:#339933" value="">==练武==
<option style="color:#339933" value="/打坐$ ">打坐练功
<option style="color:#339933" value="/闭目$ ">闭目养神
<option style="color:#339933" value="/练武$ ">武馆习武
<option style="color:#339933" value="/休身养性$ ">休身养性
<option style="color:#339933" value="/传授$ 100">传授内力
<option style="color:#FF6699" value="">==婚姻==
<option style="color:#FF6699" value="/求婚$ 真爱表白">示爱求婚
<option style="color:#FF6699" value="/情人$ 作我的情人好吗？">世纪情人
<option style="color:#FF6699" value="/分手$ 你不在是我的情人…">爱已褪色
<option style="color:#339933" value="/查看公告$ ">查看公告
<option style="color:#ff6699" value="/邀请跳舞$ ">邀请跳舞
<option style="color:#FF6699" value="/主题$ ">千里传音
<option style="color:#FF6699" value="/心动$ ">心动感觉
<option style="color:#FF6699" value="/怒吼$ ">狂狮怒吼
<option style="color:#FF6699" value="/心跳$ ">心跳心语
<option style="color:#333333" value="">==偷袭==
<option style="color:#333333" value="/单挑$ 看你拽还是我拽">不服单挑 
<option style="color:#333333" value="/吸星大法$ ">吸取内力
<option style="color:#333333" value="/经验$ 100">传授经验
<option style="color:#333333" value="/偷钱$ ">偷取钱财
<option style="color:#333333" value="/乾坤一掷$ 100000">乾坤一掷 
<option style="color:#333333" value="/下毒$ 毒药名">偷偷下毒
<option style="color:#333333" value="/投掷$ 暗器名">投掷暗器
<option style="color:#333333" value="/发招$ 招式名">发招攻击 
<option style="color:#996633" value="">==职业==
<option style="color:#996633" value="/拐骗少女">拐卖人口
<option style="color:#996633" value="/治病$ ">妙手回春
<option style="color:#996633" value="/教武$ ">教导武功
<option style="color:#996633" value="/求签$ ">求神问卜
<option style="color:#996633" value="/倒夜香$ ">倒夜香师
<option style="color:#996633" value="/出家$ 出家后，谁也不可以打你，你也不能打人。">出家修行
<option style="color:#996633" value="/还俗$">修成还俗
<option style="color:#996633" value="/唸经$">普渡众生
<option style="color:#996633" value="/如沐春风$">如沐春风
<option style="color:#996633" value="/加入天网$">加入天网
<option style="color:#996633" value="/离开天网$">脱离天网
<option style="color:#993333" value="/离开天网$">==掌门==
<option value="/招收弟子$">招收弟子
<option value="/逐出弟子$">逐出弟子
<option value="/册封$ 身份名">册封身份
<option value="/册封$ 长老">册封长老
<option value="/册封$ 护法">册封护法
<option value="/册封$ 堂主">册封堂主
<option value="/册封$ 弟子">取消册封
<option value="/掌门发放$">掌门发放
<option style="color:blue" value="/哑穴$ 请写明原因 ">教训弟子
<option style="color:blue" value="/奖励$ 100 ">奖励弟子
<option style="color:red" value="/掌门令$ 命令 ">掌 门 令
<option style="color:red" value="/本派罚款$ 1000 ">本派罚款
<option style="color:red" value="/本派刑法$ 1000 ">本派刑法
<option style="color:blue" value="/哑穴$ 请写明原因 ">哑穴操作
<option style="color:blue" value="/点穴$ 请写明原因 ">点穴操作
<option style="color:blue" value="/解穴$ 用户名">解穴操作
<option style="color:blue" value="/警告$ 请写明原因 ">警告操作
<option style="color:blue" value="/逮捕$ 请写明原因 ">逮捕操作
<option style="color:blue" value="/坐牢$ 请写明原因 ">坐牢操作
<option style="color:blue" value="/查ip$ ">查看IP
<option style="color:blue" value="/踢人$ 请写明原因 ">让你捣乱
<option style="color:blue" value="/炸弹$">吃 炸 弹
<option style="color:blue" value="/监禁$ 请写明原因 ">监禁操作
<option style="color:blue" value="/释放$ 用户名">解除监禁
<option style="color:red" value="/查看物品$">查看物品
<option style="color:red" value="/今日话题$ 话题内容……">今日话题
<option style="color:red" value="/放大$ ">放大文字
<option style="color:red" value="/状态$ ">查看状态
<option style="color:red" value="/公告$ ">站长公告
<option style="color:red" value="<script>parent.f3.location.href='#';</script>所有玩家全部存点!">全部存点
<option style="color:red" value="/清屏$ 操作">站长清屏
<option style="color:red" value="/监禁ip$ 某一人捣乱，监禁所有ip相同用户！">&nbsp;监禁ip&nbsp;
<option style="color:red" value="/封锁ip$ 原因：在聊天室中多次捣乱！">&nbsp;封锁ip&nbsp;
<option style="color:red" value="/通缉$ 通缉犯名字">通缉人犯
<option style="color:red" value="/解除$ 通缉犯名字">取消通辑
<option style="color:red" value="/站长发放$">站长发放
<option style="color:red" value="/斩首$ 原因 ">斩首人犯
</select>
      <select name='select' style='font-size:12px; background-color:#E4CCAC'>
        <option value="" selected>法力轻功</option>
        <option style="color:green" value="">==私人==</option>
        <option style="color:green" value="/传送法力$ 100">传送法力</option>
        <option style="color:green" value="/存法力$ 100">存放法力</option>
        <option style="color:green" value="/取法力$ 100">取回法力</option>
        <option style="color:green" value="/寻水晶球$ ">寻找水晶</option>
        <option style="color:green" value="/寻找法器$ ">寻找法器</option>
        <option style="color:green" value="/宝物术$ ">寻找宝物</option>
        <option style="color:green" value="/点金术$ 100">点金术</option>
        <option style="color:green" value="/解毒术$ 100">解毒术</option>
        <option style="color:green" value="/平安术$ 100">平安术</option>
        <option style="color:green" value="/偷窃术$ 100">偷窃术</option>
        <option style="color:green" value="/摇钱术$ 100">摇钱术</option>
        <option style="color:green" value="/蓝玛瑙$ 100">蓝玛瑙</option>
        <option style="color:green" value="/帅哥令$ ">帅哥令</option>
        <option style="color:green" value="/美人令$ ">美人令</option>
        <option style="color:green" value="/多情环$ ">多情环</option>
        <option style="color:green" value="/绝命钩$ ">绝命钩</option>
        <option style="color:red" value="">==使展==</option>
        <option style="color:red" value="/乞讨神术$ ">乞讨神术</option>
        <option style="color:red" value="/布施术$ ">布施天下</option>
        <option style="color:red" value="/魅惑人间$ ">魅惑人间</option>
        <option style="color:red" value="/魔界咒语$ ">魔界咒语</option>
        <option style="color:red" value="/没收法器$ ">没收法器</option>
        <option style="color:red" value="/迷魂大法$ ">迷魂大法</option>
        <option style="color:red" value="/配制令牌$ ">配制令牌</option>
        <option style="color:blue" value="">==攻击==</option>
        <option style="color:blue" value="/魔幻水晶$ ">魔幻水晶</option>
        <option style="color:blue" value="/使出$ 九阳神功 ">九阳神功</option>
        <option style="color:blue" value="/使出$ 天堂令 ">天堂令 </option>
        <option style="color:blue" value="/使出$ 玄冥棒 ">玄冥棒 </option>
        <option style="color:blue" value="/使出$ 盗取令 ">盗取令 </option>
        <option style="color:blue" value="/使出$ 七伤拳 ">七伤拳 </option>
        <option style="color:blue" value="/使出$ 圣火令 ">圣火令 </option>
        <option style="color:blue" value="/使出$ 铁笔银勾 ">铁笔银勾 </option>
        <option style="color:#FF6699" value="">==字型==</option>
        <option style="color:#FF6699" value="/粗体字$ ">粗体字</option>
        <option style="color:#FF6699" value="/飞舞字$ ">飞舞字</option>
        <option style="color:#FF6699" value="/按钮字$ ">按钮字</option>
        <option style="color:#FF6699" value="/滚动按钮$ ">滚动按钮</option>
        <option style="color:#FF6699" value="/上下按钮$ ">上下按钮</option>
        <option style="color:#FF6699" value="/爱神$ ">爱神传音</option>
        <option style="color:#FF6699" value="/颠倒$ ">颠倒字体</option>
        <option style="color:#996633" value="">==私人==</option>
        <option style="color:#996633" value="/传授轻功$ 100">传授轻功</option>
        <option style="color:#996633" value="/轻功暂存$ 100">轻功暂存</option>
        <option style="color:#996633" value="/提出轻功$ 100">提出轻功</option>
        <option style="color:#000000" value="">==偷窃==</option>
        <option style="color:#000000" value="/讨取轻功$ ">讨取轻功</option>
        <option style="color:#000000" value="/寻找秘笈$ ">寻找秘笈</option>
      </select> 
      <select name='dubo' onchange="rc(this.value);document.af.dubo.options[0].selected=true;" style='font-size:12px; background-color:#E4CCAC'>                                                                          
<option value="" selected>&nbsp;赌博&nbsp;
<option style="color:blue" value="/江湖赛马">赛&nbsp;&nbsp;马
<option style="color:blue" value="/双人赌博$ 1">赌石头
<option style="color:blue" value="/双人赌博$ 2">赌剪子
<option style="color:blue" value="/双人赌博$ 3">赌博布
<option style="color:red" value="/对赌$10">赌金币
<option style="color:red" value="/赌豆$100">赌豆点
<option style="color:red" value="/赌法力$100">赌法力
<option style="color:red" value="/赌轻功$100">赌轻功
<option style="color:red" value="/幸运$ 幸运数字1-1000">风&nbsp;&nbsp;采
<option style="color:red" value="/赌博$ 消息">赌消息
<option style="color:blue" value="/赌博$ 银作庄">银作庄
<option style="color:blue" value="/下注$ 银&大&10000">银押大
<option style="color:blue" value="/下注$ 银&小&10000">银押小
<option style="color:blue" value="/下注$ 银&双&10000">银押双
<option style="color:blue" value="/下注$ 银&单&10000">银押单
<option style="color:green" value="/赌博$ 点作庄">点作庄
<option style="color:green" value="/下注$ 点&大&100">点押大
<option style="color:green" value="/下注$ 点&小&100">点押小
<option style="color:green" value="/下注$ 点&双&100">点押双
<option style="color:green" value="/下注$ 点&单&100">点押单
  </select><select name='tu' style='font-size:12px; background-color:#E4CCAC'>                         
<option value="" selected>图片功能</option> 
<option value="[IMG]xl.gif[/IMG]">笑脸</option> 
<option value="[IMG]bm.gif[/IMG]">不满</option> 
<option value="[IMG]yy.gif[/IMG]">生气</option> 
<option value="[IMG]gl.gif[/IMG]">鬼脸</option> 
<option value="[IMG]st.gif[/IMG]">舌头</option> 
<option value="[IMG]qc.gif[/IMG]">汽车</option> 
<option value="[IMG]kl.gif[/IMG]">骷髅</option> 
<option value="[IMG]ml.gif[/IMG]">美女</option> 
<option value="[IMG]dg.gif[/IMG]">蛋糕</option> 
<option value="[IMG]qt.gif[/IMG]">拳头</option> 
<option value="[IMG]zd.gif[/IMG]">炸弹</option> 
<option value="[IMG]yy.gif[/IMG]">咬牙</option> 
<option value="[IMG]xiong.gif[/IMG]">小熊</option> 
<option value="[IMG]mg.gif[/IMG]">玫瑰</option> 
<option value="[IMG]ws.gif[/IMG]">握手</option> 
<option value="[IMG]dh.gif[/IMG]">电话</option> 
</select><select name='addsays' style='font-size:12px; background-color:#E4CCAC'>                                                                        
<option value="对" selected>表情</option> 
<option value="微微笑对">微笑</option> 
<option value="温柔地对">温柔</option> 
<option value="红着脸对">脸红</option> 
<option value="摇头晃脑得意地对">得意</option> 
<option value="哈！哈！哈！笑着对">大笑</option> 
<option value="神秘兮兮地对">神秘</option> 
<option value="战战兢兢地对">战兢</option> 
<option value="毛手毛脚地对">毛手</option> 
<option value="嘟着嘴地对">嘟嘴</option> 
<option value="慢条斯理地对">慢条</option> 
<option value="同情地对">同情</option> 
<option value="幸灾乐祸地对">乐祸</option> 
<option value="快要哭地对">快哭</option> 
<option value="哭着对">哭</option> 
<option value="拳打脚踢地对">拳打</option> 
<option value="不怀好意地对">坏意</option> 
<option value="遗憾地对">遗憾</option> 
<option value="瞪大了眼睛，很诧异地对">诧异</option> 
<option value="幸福地对">幸福</option> 
<option value="翻箱倒柜地对">翻箱</option> 
<option value="悲痛地对">悲痛</option> 
<option value="正义凛然地对">正义</option> 
<option value="严肃地对">严肃</option> 
<option value="生气地对">生气</option> 
<option value="大声地对">大声</option> 
<option value="傻乎乎地对">傻</option> 
<option value="很满足地对">满足</option> 
<option value="手足无措地对">无措</option> 
<option value="很无辜地对">无辜</option> 
<option value="喃喃自语地对">自语</option> 
<option value="恶狠狠地瞪着眼对">瞪眼</option> 
<option value="快要吐地对">想吐</option> 
<option value="无精打采地对">无采</option> 
<option value="依依不舍地对">不舍</option> 
<option value="口吐白沫对">白沫</option> 
<option value="有气无力地对">无力</option> 
<option value="无可奈何地对">无奈</option> 
<option value="可怜兮兮地对">可怜</option> 
</select>
<br>
      聊天和格斗所使用的大部分功能都集中在这里了，你试试便知用途，这里就不繁述了。 <br>

	   
      <hr noshade size="1" color=009900 align="left">
	  <input type=checkbox name='towhoway' value='1' accesskey='s'>
      <a href='#' onMouseOver="window.status='选中本项后，说的话只有你和对方才能看到（即使版主也看不到）。';return true" onMouseOut="window.status='';return true" title="悄悄话儿悄悄说">私聊</a> 
      如果你说的话只想让对方一个人看到，那务必在此打上勾喔，要不然所有人都可以看到啦 ^_^
<br><br>
			<input type="checkbox" name="addvalues">
      <a href="#" title='打开后,系统定时自己发言,自己存点,不在线一样泡！'>泡点</a> 
      如果你有事或要休息而不能再聊天<b>，那只要在这里打上勾，一样泡点！</b><br>
泡点系统可是会送不少东西哦，比如银两、武功内力等等，当然最重要的是等级升的更快啦 ^_^
<br><br>
      <a class=blue href="#" title="来答字花吧答对有银两可赚喔！">字花</a> 
      如果系统或官府的人出题了，<b>要迅速点这里答题啦，答对有奖的！</b> <br><br>
      <a href=lg.asp title="在练功的时间，你与别人是不能打架的！" target="d">出/闭关</a> 如果有人打你，而你的武功、内力、体力等又不如人的话，<b>点这里先避避吧，君子报仇，数日不晚</b></td>
  </tr>
  <tr> 
    <td bgcolor="#FFCC00"> <font color="#FF0000" size="2"><b>还有些功能没有详细介绍，大家慢慢体会吧，祝大家：早日成就高手一族！</b></font><font size="2">&nbsp;
        </font>
</td>
  </tr>
  <tr> 
    <td width="100%" align="left" style="text-indent: 0"> <p align="center"> 
      <p align="center"><a href="javascript:window.close()""><font size="4" color="#009933"><strong>好了，我明白了，咱们开始吧！</strong></font> </a> 
    </td>
  </tr>
</table>
</body>
</html>
