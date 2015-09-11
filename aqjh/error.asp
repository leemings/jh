<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>出错提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="chat/readonly/style.css">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body background=bg.gif leftmargin="0" topmargin="0" onLoad="MM_preloadImages('jhimg/err_close2.gif','jhimg/err_but2.gif')">
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<table align=center background=jhimg/err_bg1.gif border=0 
      cellpadding=0 cellspacing=0 class=body width=400>
  <tbody> 
  <tr> 
    <td height=23 width=25><img border=0 height=23 
            src="jhimg/err1.gif" width=24></td>
    <td height=23 width=340>&nbsp;<font color=#000000 
            face="Arial, Helvetica, sans-serif">ERROR - 出错啦！</font></td>
    <td align=right height=23 valign=baseline width="35"><a 
            href="javascript:window.close()" onMouseOut=MM_swapImgRestore() 
            onMouseOver="MM_swapImage('close','','jhimg/err_close2.gif',1)"><img 
            border=0 height=18 name=close src="jhimg/err_close1.gif" 
            width=15></a><img border=0 height=23 name=errr1_c4 
            src="jhimg/err2.gif" width=5></td>
  </tr>
  </tbody> 
</table>
<table border="0" width="400" cellpadding="0" cellspacing="0" align="center" background="JHIMG/err_bg.gif">
  <tr> 
    <td width="98" align="center" valign="top"><img src="Error.gif"></td>
    <td width="302"> 
      <%Select Case id
Case "000"%>
      <p>　　对不起，程序所在目录不是虚拟目录，或设置有错误，Global.asa 没有被执行。本程序需要虚拟目录的支持！如果非虚拟目录请将global.asa复制到根目录上去！</p>
      <%Case "002"%>
      <p>　　对不起，本程序不能在这台服务器上运行，如果你是本程序的合法用户，请立即与作者联系。<br>
        信箱：<a href="mailto:119yejin@163.com">119yejin@163.com</a><br>
      </p>
      <%Case "003"%>
      <p>　　数据库尚未打开，可能是站长正在整理压缩数据库，请稍候再来。</p>
      <%Case "010"%>
      <p>　　对不起，本程序专门针对 Internet Explorer 4.0 以上版本的浏览器制作，请使用 IE 浏览器来访问。</p>
      <%Case "011"%>
      <p>　　对不起，请勿通过代理服务器来访问本站。</p>
      <%Case "20"%>
      <p>　　数字5-1-7-2-6-8-0-5不正确！</p>
      <%Case "50"%>
      <p>　　您的oicq格式不正确，oicq为：5-10位的数字！</p>
      <%Case "51"%>
      <p>　　您的Email格式不正确，Email为找回密码用，非常重要！</p>
      <%Case "52"%>
      <p>　　为了您的安全，密码请使用6位以上的字符，可以用字母、数字。</p>
      <%Case "53"%>
      <p>　　请用中文输入数据！</p>
      <%Case "54"%>
      <p>　　系统禁止了“'”,“,”,“"”,“or”等对系统有破坏性的字符！</p>
      <%Case "55"%>
      <p>　　系统为了防止黑客使用探测软件所以请60秒后再测试新用户名！</p>
      <%Case "56"%>
      <p>　　用户名，oicq号，Email，密码，提示都不能为空！</p>
      <%Case "57"%>
      <p>　　介绍人的姓名不能与你自己的姓名相同！</p>
      <%Case "58"%>
      <p>　　介绍人并不存在，输入错误！</p>
      <%Case "59"%>
      <p>　　注册人与介绍人为同一ip,系统禁止注册！</p>
      <%Case "60"%>
      <p>　　介绍人含有非法字符！只能使用中文（不能带空格及其它符号）。</p>
      <%Case "61"%>
      <p>　　你不能操作，请退出江湖聊天室再进行！</p>
      <%Case "62"%>
      <p>　　新用户名已经存在！</p>
      <%Case "63"%>
      <p>　　原用户名已经不存在或密码不正确！</p>
      <%Case "64"%>
      <p>　　你是官府中人或者你是门派掌门，禁止改名！</p>
      <%Case "65"%>
      <p>　　你是江湖会员，为了方便管理理，禁止改名！</p>
      <%Case "66"%>
      <p>　　提示与答案不能相同！</p>
      <%Case "67"%>
      <p>　　因为最后注册人与您的IP相同，为防止作弊请您等上120秒再注册！(同一网吧注册时也请您等待!)</p>
      <%Case "68"%>
      <p>　　所输入id并不存在！</p>
      <%Case "69"%>
      <p>　　所输入生日不正确，我们无法查找！</p>
      <%Case "70"%>
      <p>　　所输入答案不正确，不能取回！</p>
      <%Case "71"%>
      <p>　　你不能操作，请退出江湖聊天室再进行！</p>
      <%Case "72"%>
      <p>　　你想作什么？呵。。黑老大！</p>
      <%Case "100"%>
      <p>　　欢迎您的光临，只是站长已经关闭了聊天室的登录功能，现在不能进行登录，请稍后再来。</p>
      <%Case "101"%>
      <p>　　欢迎您的光临，聊天室现在已满，站长限制最多同时在线人数为 <font color=red><%=Application("aqjh_chat_maxpeople")%></font> 
        人，请稍后再来。</p>
      <%Case "102"%>
      <p>　　站长禁止新用户名登录，请稍后再来。</p>
      <%Case "110"%>
      <p>　　您现在的 IP：<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 
        被封锁 <%=Application("aqjh_iplocktime")%> 分钟，不能进入聊天室。<br>
        　　离自动解锁时间还有：<font color=red><%=ABS(int(Application("aqjh_iplocktime"))-int(Datediff("s",Request.QueryString("lockdate"),now())/60))%></font> 
        分钟。</p>
      <%Case "111"%>
      <p>　　您现在的 IP：<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 
        被<font color=red>【永久】</font>封锁！不能进入聊天室。请与站长联系。</p>
      <%Case "120"%>
      <p>　　用户名含有非法字符！只能使用中文（不能带空格及其它符号,请改名）。</p>
      <%Case "121"%>
      <p>　　密码含有非法字符！只能使用英文字母和数字（不能带空格）。</p>
      <%Case "122"%>
      <p>　　称谓含有非法字符！只能使用中文、英文字母和数字（不能带空格）。</p>
      <%Case "123"%>
      <p>　　新密码含有非法字符！只能使用英文字母和数字（不能带空格）。</p>
      <%Case "124"%>
      <p>　　确认密码含有非法字符！只能使用英文字母和数字（不能带空格）。</p>
      <%Case "125"%>
      <p>　　用户名的长度超过 10 个字符（一个汉字占两个字符）。</p>
      <%Case "126"%>
      <p>　　称谓的长度超过 4 个字符（一个汉字占两个字符）。</p>
      <%Case "127"%>
      <p>　　用户名不能为空。</p>
      <%Case "128"%>
      <p>　　密码不能为空。</p>
      <%Case "129"%>
      <p>　　密码不能和用户名相同。</p>
      <%Case "130"%>
      <p>　　对不起，该用户名为系统所保留，您不能使用这个名字注册或登录。</p>
      <%Case "131"%>
      <p>　　对不起，该用户名含有不雅字眼，您不能使用这个名字注册或登录。</p>
      <%Case "140"%>
      <p>　　对不起，该用户名正在使用中，不能完成您所需要的操作。<br>
        　　如果您是首次使用该用户名登录，则可能是该用户名已经被其他人抢注了，您只能换用其他名字登录。<br>
        　　如果您以前曾经使用过该用户名登录成功，则可能是您的用户名被其他人盗用。<br>
        另一种可能是：您没有正常退出聊天室（如：掉线、超时与服务器断开连接等），请使用掉线处理！导致用户名被卡在聊天室中，则请您关闭所有浏览器，再重新打开，并等十分钟后再来登录。如果实在不行，请到留言簿给版主留言，请版主为您解决。</p>
      <%Case "141"%>
      <p>　　对不起，密码错误（注意：密码区分大小写），不能完成您所需要的操作。<br>
        　　如果您是首次使用该用户名登录，则可能是该用户名已经被其他人抢注了，您只能换用其他名字登录。<br>
        　　如果您以前曾经使用过该用户名登录成功，则可能是您的用户名被其他人盗用，并且盗用者更改了密码。如果是这种情况，请到留言簿给版主留言，请版主为您解决。</p>
      <%Case "142"%>
      <p>　　对不起，该用户名被禁用！请到留言簿给版主留言，请版主为您解决。</p>
      <%Case "143"%>
      <p>　　对不起，该用户名被踢出聊天室，原因:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,5 
        分钟内不能使用该用户名登录。还有 <font color=red><%=ABS(300-Datediff("s",Request.QueryString("lastkick"),now()))%></font> 
        秒。</p>
      <%Case "150"%>
      <p>　　对不起，该用户名尚未被注册，当然不能修改密码了。</p>
      <%Case "151"%>
      <p>　　对不起，您没有指定“新密码”，怎么修改密码呢？</p>
      <%Case "152"%>
      <p>　　“新密码”与旧密码完全相同，用不着再修改密码啦！</p>
      <%Case "153"%>
      <p>　　“新密码”不能与用户名相同！</p>
      <%Case "160"%>
      <p>　　对不起，该用户名尚未被注册，当然不能“自杀”了。</p>
      <%Case "161"%>
      <p>　　“确认密码”为空，不能执行自杀操作。</p>
      <%Case "162"%>
      <p>　　“确认密码”与“密码”不一致，不能执行自杀操作。</p>
      <%Case "163"%>
      <p>　　该用户名被永久保留，不能执行自杀操作。</p>
      <%Case "164"%>
      <p>　　输入有误，不能执行自杀操作。</p>
      <%Case "165"%>
      <p>　　有没有搞错，你想搞谋杀啊！</p>
      <%Case "166"%>
      <p>　　两次密码必需相同才能注册！</p>
      <%Case "167"%>
      <p>　　用户名己经被注册，请选用其他用户名！</p>
      <%Case "001"%>
      <p>　　该程序执行了非法操作，请立即停止使用，否则可能引起不可预测的后果，请立即与作者取得联系。<br>
        信箱：<a href="mailto:qq9yejin@163.com">119yejin@163.com</a><br>
      </p>
      <%Case "200"%>
      <p>　　没有此类可供显示的留言数据，不能查看。</p>
      <%Case "201"%>
      <p>　　没有此类可供显示的数据，不能查看。</p>
      <%Case "210"%>
      <p>　　你尚未登录，或已经超时断开了连接，请重新登录。</p>
      <%Case "220"%>
      <p>　　“来自何方”含有非法字符，也不能使用HTML语法！</p>
      <%Case "221"%>
      <p>　　“E-Mail”地址含有非法字符，也不能使用HTML语法！</p>
      <%Case "222"%>
      <p>　　“主页地址”含有非法字符，也不能使用HTML语法！</p>
      <%Case "223"%>
      <p>　　“来自何方”长度超过30字符，1个汉字占2字符。</p>
      <%Case "224"%>
      <p>　　“E-mail”长度超过50字符。</p>
      <%Case "225"%>
      <p>　　“主页地址”长度超过50字符 。</p>
      <%Case "226"%>
      <p>　　“个人简介”长度超过200字符 。</p>
      <%Case "227"%>
      <p>　　“E-mail”地址格式有误。</p>
      <%Case "228"%>
      <p>　　“主页地址”格式有误 。</p>
      <%Case "229"%>
      <p>　　用户名不存在，不能修改个人信息。</p>
      <%Case "230"%>
      <p>　　请输入要查询的用户名。</p>
      <%Case "231"%>
      <p>　　用户名：<font color=red><%=Request.QueryString("name")%></font> 不存在。</p>
      <%Case "240"%>
      <p>　　关键词为空，不能搜索。</p>
      <%Case "250"%>
      <p>　　所有“红色”的项目均为必填项目，请填写完整。</p>
      <%Case "251"%>
      <p>　　用户名：<font color=red><%=Request.QueryString("name")%></font> 不存在，不能使用“悄悄话”方式。</p>
      <%Case "252"%>
      <p>　　密码为空，不能使用用户名：<font color=red><%=Request.QueryString("name")%></font> 
        进行留言。</p>
      <%Case "253"%>
      <p>　　密码错误，不能使用用户名：<font color=red><%=Request.QueryString("name")%></font> 
        进行留言。</p>
      <%Case "254"%>
      <p>　　“写给”的长度超过20个字符（1个汉字=2个字符）。</p>
      <%Case "255"%>
      <p>　　“主题”的长度超过40个字符（1个汉字=2个字符）。</p>
      <%Case "256"%>
      <p>　　“内容”的长度超过1024个字。</p>
      <%Case "257"%>
      <p>　　“姓名”的长度超过20个字符（1个汉字=2个字符）。</p>
      <%Case "258"%>
      <p>　　“地址”的长度超过20个字符（1个汉字=2个字符）。</p>
      <%Case "259"%>
      <p>　　“信箱”的长度超过50个字符（1个汉字=2个字符）。</p>
      <%Case "260"%>
      <p>　　“主页名称”的长度超过24个字符（1个汉字=2个字符）。</p>
      <%Case "261"%>
      <p>　　“主页地址”的长度超过50个字符。</p>
      <%Case "262"%>
      <p>　　“信箱”输入有误，请重新输入。</p>
      <%Case "263"%>
      <p>　　“主页地址”的格式有误。</p>
      <%Case "264"%>
      <p>　　留言“内容”中不能出现连续的空白行。</p>
      <%Case "265"%>
      <p>　　不能重复粘贴相同的留言。</p>
      <%Case "266"%>
      <p>　　“主题”不能包含半角的“"”“'”引号。</p>
      <%Case "300"%>
      <p>　　该用户名已经在聊天室中，不能重复进入,请使用掉线管理再进入江湖！。</p>
      <%Case "301"%>
      <p>　　不能以“江湖管理员”的名称进入聊天室中，请重回登录页面换名登录。</p>
      <%Case "400"%>
      <p>　　对象：<font color=red><%=Request.QueryString("name")%></font> 不在聊天室中，不能发送消息。</p>
      <%Case "401"%>
      <p>　　消息的长度必须小于 1024 个字符。</p>
      <%Case "402"%>
      <p>　　消息中不能出现连续的空白行。</p>
      <%Case "403"%>
      <p>　　消息不能为空。</p>
      <%Case "410"%>
      <p>　　对象：<font color=red><%=Request.QueryString("name")%></font> 不在聊天室中，不能为其点歌。</p>
      <%Case "420"%>
      <p>　　由于你被抓进牢里，原因:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,等着保释你吧！不要再干坏事了！</p>
      <%Case "421"%>
      <p>　　你被人打死了，原因:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>，还是到<a href="yamen/disp.asp">阎王殿</a>来吧！</p>
      <%Case "422"%>
      <p>　　由于你被抓进牢里，原因:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,等10分钟后被释放吧，不要再干坏事了！</p>
      <%Case "423"%>
      <p>　　由于你未注册、者账号被盗改名首或者没有及时还贷款被杀！！请重新<a href='yamen/joinjh.asp'>注册</a>你的账号吧！</p>
      <%Case "424"%>
      <p>　　由于你被人点穴！一时间还没有醒来过了！</p>
      <%Case "425"%>
      <p>　　你不是官府的人，你没有权利访问这里！</p>
      <%Case "426"%>
      <p>　　我不是长老，你没有权利访问问这里！</p>
      <%Case "427"%>
      <p>　　你不是江湖中人，不能结婚！</p>
      <%Case "428"%>
      <p>　　怎么回事？不会是同性恋吧！</p>
      <%Case "429"%>
      <p>　　你来慢一步了……</p>
      <%Case "430"%>
      <p>　　不能登记！对方的江湖等级不够5级，配偶不为空或者你俩性别相同！！</p>
      <%Case "431"%>
      <p>　　不能登记！因为你已经登记了请3天后再登记！</p>
      <%Case "432"%>
      <p>　　你的密码不能为空！</p>
      <%Case "433"%>
      <p>　　伴侣的名字、聘礼大于10万，信息不能为空！</p>
      <%Case "4333"%>
      <p>　　你的等级不够，你想作什么，申请配偶需要3级以上！</p>
      <%Case "4334"%>
      <p>　　她(他)的等级不够，你想作什么，申请配偶需要3级以上！</p>
      <%Case "4335"%>
      <p>　　江湖上找不到你想取的人！</p>
      <%Case "4336"%>
      <p>　　call!你想作什么？这里不欢迎你，滚！</p>
      <%Case "434"%>
      <p>　　你想自己取自己！！古往今来都没有这回事！</p>
      <%Case "435"%>
      <p>　　开什么玩笑，你们俩性别一样啊！江湖里是不允许同性恋的。</p>
      <%Case "436"%>
      <p>　　只听里面传来一片尖叫，你慌慌张张的逃了出来。</p>
      <%Case "437"%>
      <p>　　你好象没有带那么多钱了吧？</p>
      <%Case "438"%>
      <p>　　对不起，你今天已经很干净了，温泉浴每天只可以洗一次。</p>
      <%Case "439"%>
      <p>　　你没有登录或你不是官府中人！你不能进入管理区。</p>
      <%Case "440"%>
      <p>　　SORRY！您没有<a href=index.asp>登录</a></p>
      <%Case "441"%>
      <p>　　未找到掌门的资料，掌门未定！</p>
      <%Case "442"%>
      <p>　　错误：以上项目中门派、适合必须填写，适合必须为男、女、所有！</p>
      <%Case "443"%>
      <p>　　错误：观看概率超出范围，长老，看好再填啊！</p>
      <%Case "444"%>
      <p>　　错误：帮派名称已经存在！</p>
      <%Case "445"%>
      <p>　　错误：帮派名称必须填写！！</p>
      <%Case "446"%>
      <p>　　未找到掌门的资料，掌门已经被打死了？？？.....不会吧，这么菜的掌门！</p>
      <%Case "447"%>
      <p>　　该门派已不存在！请刷新页面！</p>
      <%Case "448"%>
      <p>　　官府不允许改名换姓！</p>
      <%Case "449"%>
      <p>　　错误：你不是江湖人,请不要误闯禁区！</p>
      <%Case "449"%>
      <p>　　未找到掌门的资料，掌门未定！</p>
      <%Case "450"%>
      <p>　　你帮今日己领了一次银两了，金库不再支付你的开销，省着点吧！</p>
      <%Case "451"%>
      <p>　　错误：你不是掌门,请不要误闯禁区！</p>
      <%Case "452"%>
      <p>　　错误：你是官府的人，不可以改名！</p>
      <%Case "453"%>
      <p>　　错误：你不是江湖中人，请不要随便乱闯！</p>
      <%Case "454"%>
      <p>　　错误：你不是江湖中人，我们不收你这种打短工的小人物！</p>
      <%Case "455"%>
      <p>　　错误：你无门无派，不许进入本帮禁地！</p>
      <%Case "456"%>
      <p>　　错误：这个漏洞己堵上，请不要再试了！</p>
      <%Case "457"%>
      <p>　　错误：没钱也想改名？？先赚点银子再来改名吧！</p>
      <%Case "458"%>
      <p>　　错误：你的银两不够，我们不接受你的离婚登记！</p>
      <%Case "459"%>
      <p>　　错误：操作不成功!!请返回！</p>
      <%Case "460"%>
      <p>　　错误：你等级太低了点吧！要是创派后弟子都养不活吧！</p>
      <%Case "461"%>
      <p>　　错误：你好像有门派了吧！有多大能奈，还想建多少个派呀！</p>
      <%Case "462"%>
      <p>　　错误：你的银两买不起这件物品！</p>
      <%Case "463"%>
      <p>　　错误：你没有寄售的物品在本集市！</p>
      <%Case "464"%>
      <p>　　错误：输入的内容出错，请确认您输入的号码为数字！</p>
      <%Case "465"%>
      <p>　　对不起，由于您行为不端，道德太低，罚你到<a href="xg.asp">思过壁</a>“面壁思过”。</p>
      <%Case "466"%>
      <p>　　对不起，因钱庄小本经营，您的现金数目太多，请下次再来！</p>
      <%Case "467"%>
      <p>　　江湖高手，你武功超过150000上限，系统已经自动将其降至1499999，请不要再提升武功了，你已经是江湖一等一高手了，请<a href="index.htm">重新登录</a>即可！</p>
      <%Case "468"%>
      <p>　　江湖高手，你体力超过1亿上限，系统已经自动将其降至99000000，请不要再提升体力了，你已经是江湖一等一高手了，请<a href="index.htm">重新登录</a>即可！</p>
      <%Case "469"%>
      <p>　　你没有这种类型的装备，请到集市购买相关装备</p>
      <%Case "470"%>
      <p>　　你没有任何物品。</p>
      <%Case "471"%>
      <p>　　不能设置负数参数。</p>
      <%Case "472"%>
      <p>　　由于你在休息中，现在还没有醒来，请过一段时间再来吧！</p>
      <%Case "473"%>
      <p>　　由于你违反江湖规则，已被终生监禁，此账号不能使用了！</p>
      <%Case "474"%>
      <p>　　由于你人品太差，道德小于-10000，江湖不欢迎你，请去采冰、采矿加钱去孤儿院买道德去吧！！</p>
      <%Case "475"%>
      <p>　　对不起，你输入的数据有一项是空值！</p>
      <%Case "476"%>
      <p>　　对不起,官府目前并未出题给你答！</p>
      <%Case "477"%>
      <p>　　对不起,为了公平,官府的人不能参于此活动！</p>
      <%Case "478"%>
      <p>　　呵...黑客老大,这里不欢迎你过来,想玩去别人家吧!</p>
      <%Case "479"%>
      <p>　　对不起,你所提供的答案不对！你得不到奖励！</p>
      <%Case "480"%>
      <p>　　对不起,您被点穴了，原因:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,过几秒钟会穴道会自动解开,请一会再登录！</p>
      <%Case "481"%>
      <p>　　想作梦取媳妇呀？好好想想吧你！</p>
      <%Case "482"%>
      <p>　　你的等级太低了，代不了这些款！</p>
      <%Case "483"%>
      <p>　　你已经代过款了，请不要理贷款了！</p>
      <%Case "483"%>
      <p>　　因为你作恶太多，道德少于-100所以不能进入聊天室请点大教堂！</p>
      <%Case "484"%>
      <p>　　所注册名字与江湖本身的NPC人物重名, 请更换名字重新注册！</p>
      <%Case "485"%>
      <p>　　第二道密码不能和用户密码相同！</p>
      <%Case "486"%>
      <p>　　此功能只有正站长才能操作！</p>
      <%Case else%>
      <p>　　对不起，该出错类型未被登记，请联系站长解决。</p>
      <%End Select%>
    </td>
  </tr>
  <tr> 
    <td colspan="2" align="center" valign="top" height="17"> <a href="javascript:history.go(-1)" 
            onMouseOut=MM_swapImgRestore() 
            onMouseOver="MM_swapImage('back','','jhimg/err_but2.gif',1)"><img 
            border=0 height=20 name=back src="jhimg/err_but1.gif" 
            width=73></a><br>
      <img height=5 src="jhimg/err_bot.gif" 
        width=400></td>
  </tr>
</table></body></html>