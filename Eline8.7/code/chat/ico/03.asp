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
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:ffffff;text-decoration:none;}
a:hover{color:ffffff;text-decoration:none;}
</style>
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新功能说明</title>
</head>

<body bgcolor="#000000" text="#FFFFCC">

<table border="1" width="100%" bordercolor="#000000" bordercolorlight="#FFFFFF" bordercolordark="#000000">
  <tr>
    <td width="100%"><img border="0" src="ki17.gif" align="absmiddle" width="18" height="18">新功能说明-程序设计江湖游戏网-混世魔王</td>
  </tr>
  <tr>
    <td width="100%">增加功能项目</td>
  </tr>
  <tr>
    <td width="100%">法力----传送法力.存放法力.取回法力.寻找水晶.寻找法器.乞讨神术.魅惑人间.布施天下.魔界咒语.没收法器.迷魂大法.配置令牌.魔幻水晶.九阳神功.圣火令.天堂令.玄冥棒.七伤拳.盗取令.铁笔银钩</td>
  </tr>
  <tr>
    <td width="100%">轻功----传授轻功.轻功暂存.提取轻功.行走功能.寻找秘笈.讨取轻功.</td>
  </tr>
  <tr>
    <td width="100%">字型----爱神传音.颠倒字体.粗体字.按钮字.飞舞字.滚动按钮.上下按钮.</td>
  </tr>
  <tr>
    <td width="100%">职业----魔法师.武师.大夫.郎中.算命师.气功师.</td>
  </tr>
  <tr>
    <td width="100%">三国战士</td>
  </tr>
  <tr>
    <td width="100%">僵尸王</td>
  </tr>
  <tr>
    <td width="100%">魔法师可执行以下功能</td>
  </tr>
  <tr>
    <td width="100%">寻找宝物.点金术.解毒术.平安术.偷窃术.蓝玛瑙.帅哥令.美人令.摇钱术.多情环.绝命钩.</td>
  </tr>
  <tr>
    <td width="100%" align="center">往后会不断提供免费的功能给各位</td>
  </tr>
  <tr>
    <td width="100%" align="center"><a href="http://zhzx.jjedu.org/eline/" target="_blank">快乐江湖</a></td>
  </tr>
</table>

</body>

</html>
<Script >
parent.msgfrm.rows = '*,*,*';
parent.tbclu=true;
</script>