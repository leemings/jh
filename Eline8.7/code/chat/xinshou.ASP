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
<title>新手入门 - 51eline.com</title>
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
    <td width="100%" bgcolor="#FFCC00" > <font color="#008080"><b><font color="#FF0000">江湖入门</font></b></font><br>
      <font size="2" color="#008080">一个江湖的核心就是聊天室，首先请查看 <b><a href="jmhelp.asp">聊天室界面帮助</a> 
      学会怎样使用聊天室</b> <br>
      <br>
      聊天室的许多功能都在 
      <select name='command' style='font-size:12px; background-color:#E4CCAC'>                                                                          

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
      里面，点击一个命令，会生成一个命令行，如：<br>
      <br>
      <input name='sytemp' type=text style="height: 18; background-image: url('ico/wfy_i_c_g_b25.gif'); border-style: groove; border-width: 1" value="/送钱$ 1000"  size=35 maxlength=250>
      </font> <font size="2" color="#008080"> 
      <input type=submit name='subsay' value='发送' style="height: 18; background-color:E4CCAC;color:000000;border: 1 double" onmouseover="this.style.color='FF0000'" onmouseout="this.style.color='000000'">
      <br>
      按</font><font color="#008080"><font size="2"> 
      <input type=submit name='subsay' value='发送' style="height: 18; background-color:E4CCAC;color:000000;border: 1 double" onmouseover="this.style.color='FF0000'" onmouseout="this.style.color='000000'">
      就可以送出这个命令</font></font><br> <br>
      <font size="2"> 总之，多看多用，你不难会发现很多地方需要级别，内力与银子，下面再看《江湖秘笈》<br>
      </font> </td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#FFCC00" > <p><font size="2"><b><font color="#FF0000">江湖秘笈</font></b></font> 
      </p>
      <p>江湖中身份的象征就是等级了，在『E线江湖』中，等级分为管理等级与战斗等级。<br>
        <font color="#FF0000">管理等级</font>是1到10级，掌门5级，6级以上就是官府，10级为站长或代理站长。<br>
        <font color="#FF0000">战斗等级</font>是没有限制的。</p>
      <p>在大部份时候，我们说的等级都是指战斗等级 </p>
      <hr noshade size="1" color=009900 align="left"> 
      <p> 江湖中最重要的就是<font color="#FF0000">存点</font>、<font color="#FF0000">生命</font>(也就是体力)、<font color="#FF0000">内力</font>、<font color="#FF0000">银子</font><br>
        有了存点就可以升级，有了生命才能生存与运动，有了内力才能习武防身杀人，有了银子才能买到想要的东西</p>
      <p><font size="2">既然叫做“江湖”，那么您一定想快意恩仇，嗯，那么到江湖聊天室去吧。在那里您可以用本派的武功向别人发起攻击（点击聊天室右下脚的<font color="#FF0000">“动武”</font>），以达到使其降分的目的，如果对方的体力下降到-1000时，可能就要出人命了……，<br>
        但在聊天室里绝对不允许恶意的辱骂，否则官府的人可能会对你使用<font color="#FF0000">点哑穴</font>，那样您就只能发呆，逃跑慢了就只好眼看着别人对你进攻、下毒，生命值和内力一点点丧失……，<br>
        如果严重的话，官府的人会使用<font color="#FF0000">点死穴</font>把你踢出聊天室或者把你抓起来哦！关进监狱的滋味可不好受，更要命的是，这时候如果没有好朋友去救你（——要花1000两白花花的银子哦），三天以后，你就会被押赴刑场斩首示众了！！！
        </font></p>
      <p><font size="2">&nbsp;&nbsp; 这个时候，您就知道金钱和内力有多么重要了！！当您受到进攻以后，内力、生命力就会下降，如果你是一个大富翁，就赶紧去<font color="#FF0000">药铺</font>买点药品补一补，那样就能恢复了，攻击别人的时候如果您有一些称手的<font color="#FF0000">兵器</font>和比较有效的毒药，往往会事半功倍，当然，兵器铺和药铺的老板不是慈善家，会向您收一些银子，他们也要吃饭啊！！</font></p>
      <p><font size="2">&nbsp;&nbsp; 如果您没有钱了，又急着要买东西，怎么办？没关系，你可以把自己的物品拿出来<font color="#FF0000">拍卖</font>，不过得有人愿意发钱买才行啦…………</font></p>
      <p>1.最简单的方法就是泡，也就是在聊天室呆的时间越长，得到的就越多，得到些什么？点击聊天室右下角的“存点”就可以了。<br>
        另外如果结婚了有了<font color="#FF0000">伴侣</font>，你泡点的同时，伴侣的内力和银两也会增加。<br>
        <br>
        2.还有就是劳动，打工,如采冰，挖矿这些。</p>
      <p>3.再有就是不劳而获凭运气，赌了，可以在聊天室赌，也可以在外面的赌场赌，再就是<font color="#FF0000">打扑克</font>，<font color="#FF0000">打麻将</font>这些复杂一点要凭一点本事的赌。</p>
      <p>4.钱能通神，这是最基本的规律，总之有了钱，可以买药增强体力内力，可以购买武器，在<font color="#FF0000">物品（随身行囊）</font>中点击<font color="#FF0000">服用</font>，或<font color="#FF0000">装备</font>都可以<br>
        当然也可以在聊天室中用命令行使用,如投掷暗器，也可以送给别人，如送钱，送东西。</p>
      <p>5.其他的在<font color="#FF0000">闯荡(闯荡江湖)</font>这些游戏中也可以提升自已的内力或得到一些其他的东西，当然如果你不是高手或者不是手急眼快，可能会失去一些体力或其他的东西。搞不好小命也没了。</p>
      <p>6.还有一种方法就是拉关系，最好当然是加入帮派了，<font size="2">朋友多了好办事啊！有了大家的帮助，您就会感到不再孤单，当有人欺负你的时候（——你刚刚进入聊天室的时候，很可能马上变成别人练功的靶子哦！因为你现在弱不禁风，很好欺负的^0^），便会得到帮助了，您的帮会的帮主一定是一个有威望、武功高强的一代大侠，他会指导你怎样练功增加内力，怎样发贴子获得魅力，或者带领大家为了保卫自己的地盘向敌人发起进攻……，最最重要的是，帮主一般都是富翁啊！他会给你救命的银子和药品！甚至会传授你内力！当然，如果您将来的威信和武功超过帮主，他的地位……，嘿嘿！！</font></p>
      <p>7.有个一官半职，或者在门派中有点身份就可以每天领到工资了。</p>
      <p><font size="2">身上带的银子多了，有时候会遇上强盗的，您不妨把钱存到钱庄，那里是官方银行，绝对不会倒闭，利息很高的哦！！！每天1%！</font></p>
      <hr noshade size="1" color=009900 align="left"> <p>其他的一些如门派威望(在门派库金中捐钱增加威望),资质(在聊天室练习天外天静坐来提升资质),上限(使用神族的法术可以提升上限),金币(泡点增加),累了，以后再补充吧，反正这些基本上可以很快就会自已弄明白的。总不能背熟了再去玩江湖</p>
      <hr noshade size="1" color=009900 align="left"> 
<font size="2" color="#008080">大侠您闯荡江湖，腥风血雨，刀光剑影之后一定伤痕累累，还是成立一个温馨美满的小家庭吧？有个侠肝义胆、柔情似水的伴侣，江湖上的险恶就会变得不那么让人伤脑筋了……但是并不是只要您有心上人就可以结为夫妇，这个地方有个规矩：等级必须30级，老婆不是随便要的嘛 
        ^_^</font>
</td>
  </tr>
  <tr> 
    <td bgcolor="#FFCC00"> <p><font color="#FF0000" size="2"><b>江湖规矩</b></font> 
        <font size="2"><br>
        1.江湖中不允许骂人，说脏话。杀人不犯法，骂人要坐牢。<br>
        2.不要粘贴大量重复的话一次性发出，也不允许多次发送一句重复的话。<br>
        3.不允许在江湖中拉人</font> <font size="2">......</font> 最好不要在江湖中谈论其他的江湖(特别是站长禁制了的网址)。<br>
        上面这些基本上是会立即被系统自动踢出的。</p>
      <p>4<font size="2">.</font>其他的最好是不要威胁恐吓站长，在聊天室中别人至少有几百种方法让你倒霉<br>
        5<font size="2">.</font>不要试图做不该做的事，访问系统管理页面或进行破坏，『E线江湖』的这些地方都设置了炸弹保护，悠着点</p>
      <p>6<font size="2">.</font>做官府的管理员一定要尽职尽责，公正无私，每个管理员的操作都会被记录在管理日志中，供大家临督。<br>
        7<font size="2">.官府人员只管理刷屏、骂人，及违反中国有关互联网规定的对象。可以使用警告、逮捕、封锁IP的手段制止恶意捣乱的行为的发生。</font> 
        <br>
        8<font size="2">.官府人员不得参与江湖恩怨，不可以偏向任何一方，只要没有捣乱行为出现，杀人的事情一般不应过问。——滥杀新手除外，欺负新朋友的人一定要加以惩罚！</font> 
        <br>
        9<font size="2">.</font><font size="2">如果官府人员很长时间不能上网，影响了自己的工作和秩序，我们会考虑选举新的管理人员来顶替。<br>
        原则上要求每天在聊天室的时间不少于3个小时。</font></p>
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
