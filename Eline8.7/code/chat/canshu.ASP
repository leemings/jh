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
<title>江湖设定档案 - happyjh.com</title>
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
      2.</font><a href="xinshou.ASP"><font color="#FFFF00">江湖入门帮助</font></a> <font color="#ffff00">3.</font><a href="jhgj.ASP"><font color="#FFFF00">江湖规矩 
      </font></a><font color="#ffff00">4.</font><a href="canshu.asp"><font color="#FFFF00">江湖设定档案</font></a></b> 
    </td>
  </tr>

  <tr> 
    <td width="100%" align="left" bgcolor="#FFCC00" > <font color="#FF0000" size="2"><b>江湖设定</b> 
      </font> <font size="2" color="#008080">下面是当前站长在『快乐江湖』中设定的一些公开的系统参数</font><font size="2" color="#008080"><br>
      千里传音等级:20<br>
      下毒等级:20<br>
      偷钱等级:15<br>
      吸星大法:18<br>
      暗器等级:20<br>
      魔宝等级:25<br>
      发招攻击:13<br>
      传授内力:25<br>
      送钱等级:20<br>
      送法力等级:25<br>
      赠送等级:15<br>
      拍卖等级:15<br>
      心动等级:20<br>
      拜师等级:20<br>
      传经验的等级:160<br>
      转账需要的等级:15<br>
      怒吼需要等级:30<br>
      心跳需要等级:20<br>
      在线求婚:20<br>
      赌博作庄:20<br>
      情人需要等级:25 <br>
      单挑需要等级:15<br>
      求签等级:20<br>
      送花等级:12<br>
      暂离开等级:12<br>
      心情功能等级:15<br>
      标题等级:60<br>
      出家等级:15<br>
      打扑克等级:15<br>
      打麻将等级:15<br>
      寻金币等级:80<br>
      唸经需要等级:20<br>
      加入天网等级:30<br><br>
      <strong><font color="#FF0000" size="3">江湖比武时间为每小时的后30分钟</font></strong><br>
      <br>
      是否允许新人注册:允许<br>
      被封IP的时间:1200000分钟<br>
      聊天室最多不发言时间:30分钟<br>
      江湖一天的杀人数.夺宝一次算一次杀人:10<br>
      泡分时加的存点:10点/分钟<br>
      泡分时加的银两:1000两/分钟<br>
      泡分时加的武功内力:50点/分钟<br>
      <br>
      系统发钱:每小时的前10分钟 最多5000000<br>
      站长发放:<br>
      具体有:银两，最大值为1000万两;金币，最大值为5个;存点，最大值为100点; 豆点，最大值为100点<br>
      元宝，最大银两10万;体力，最高发放5千点;内力，最高发放5千点;武功，最高发放5千点，还有很多好东东哦。<br>
      <font size="2" color="#008080"><font color="#FF3300">站长常选择人多时或给难得的时间还在江湖的朋友发放</font></font><font color="#FF3300">，希望您在场哦^_^</font><br>
      <font color="#009933"> <font color="#000000">会员泡点倍数</font><font color="#008080">:</font><font color="#000000">一级会员1.2倍；二级会员1.4倍；三级会员1.6倍；四级会员1.8倍；五级会员2.5倍；泡点制会员2倍</font></font><br>
      </font> 
      <hr noshade size="1" color=009900 align="left">
      <font size="2" color="#FF0000"><strong>工资发放</strong> (请记得每天领取工资哦)</font><br>
      掌门：会员x5万+弟子数x5千+介绍人数x10000+战斗等级x15000+2万<br>
      弟子：会员等级x200000+介绍人数x10000+江湖战斗等级x15000+2万<br>
      <hr noshade size="1" color=009900 align="left">
      <font size="2" color="#FF0000"> <strong>江湖等级计算公式<br>
      </strong><font color="#008080">等级x等级x50即是存点的值，反过来看，如某人有5000点则是5000除50等于100，然后再开平方为10级，则他战斗等级为10级，由此侠客们应该要知道等级越高升的越慢。<br>
      </font></font> 
      <hr noshade size="1" color=009900 align="left">
      <font size="2" color="#FF0000"><strong>江湖开设的房间</strong><br>
      </font> <font size="2" color="#FF0000"> <font color="#008080">房间名字：快乐江湖 
      进入条件：所有(聊天)<br>
      </font></font> <font size="2" color="#008080">房间名字：高手E线 进入条件：战斗等级&gt;30(格斗)<br>
      房间名字：泡点E线 进入条件：所有(泡点)<br>
      <font size="2" color="#008080">房间名字：烂漫E线 进入条件：江湖等级&gt;50(情聊)</font> </font></td>
  </tr>
</table>
</body>
</html>
