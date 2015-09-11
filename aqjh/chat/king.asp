<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"
n=request("n")
if n="" then n="个人行李"
%>
<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<style type=text/css>
A {text-decoration: none; color:'yellow'; font-size: 9pt;}
A:hover {text-decoration: none; color:ff6600;  font-size: 9pt;}
td {text-decoration: none; color:'yellow'; font-size: 9pt;}
</style>
</head>
<body leftmargin="0" topmargin="20" bgproperties="fixed" background=bg.gif>
<table width=100%><tbody><tr><td align=center>
<br><a href=?n=个人行李>◇个人行李◇</a><br>
<a href=?n=天外有天>◇天外有天◇</a><br>
<a href=?n=江湖集市>◇江湖集市◇</a><br>
<a href=?n=门派事宜>◇门派事宜◇</a><br>
<a href=?n=官府衙门>◇官府衙门◇</a><br>
<a href=?n=会员专区>◇会员专区◇</a><br>
<a href=?n=其他功能>◇其他功能◇</a><br><br>
<%
select case n
case "个人行李"
%>
<table width=120>
<tr><td><a href=# onClick="window.open('../hcjs/jhzb/index.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">我的装备</A></td><td><a href=# onclick="window.open('../hcjs/jhzb/myhua.asp','aqjh_win','scrollbars=0,resizable=0,width=430,height=340')">我的鲜花</a></td></tr>
<tr><td><a href=# onClick="window.open('../JHMP/MP5.ASP','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">我的徒弟</a></td><td><a href=# onClick="window.open('../hcjs/box/index.asp','aqjh_win','scrollbars=1,resizable=1,width=760,height=400')">我的储物</a></td</tr>
<tr><td><a href=# onClick="window.open('../yamen/regmodi.asp','aqjh_win','scrollbars=1,resizable=0,width=480,height=470')">档案修改</a></td><td><a href=# onClick="window.open('../yamen/editpass2.asp','aqjh_win','scrollbars=1,resizable=0,width=520,height=340')">密码保护</A></td></tr>
<tr><td><a href=# onClick="window.open('../yamen/modify.asp','aqjh_win','scrollbars=1,resizable=0,width=520,height=340')">修改密码</a></td><td><a href=# onClick="window.open('../hcjs/jhmt/','aqjh_win','width=550,height=410,resizable=0,scrollbars=0')">江湖密探</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/wuguan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">江湖密室</a></td><td><a href=# onClick="window.open('../HCJS/zstt/zs.asp','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">转世投胎</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/money.asp','aqjh_win','scrollbars=0,resizable=0,width=430,height=440')">领取工资</a></td><td><a href=# onclick="window.open('../work/changeworkzhongzu.asp','aqjh_win','scrollbars=1,resizable=1,width=400,height=300')">更换种族</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/myfriend.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')">拉人记录</a></td><td><a href=# onClick="window.open('../hcjs/sendphoto/PHOTO.ASP','aqjh_win','scrollbars=1,resizable=1,width=750,height=500')">网友照片</a></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/long/long.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')">神龙商店</a></td><td><a href=# onClick="window.open('../hcjs/long/longeat.asp','aqjh_win','scrollbars=1,resizable=1,width=750,height=500')">购买龙粮</a></td></tr>
</table>
<%
case "天外有天"
%>
<table width=120>
<tr><td><a href=# onClick="window.open('../WORK/CATCH/CATCH.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">出海走私</A></td><td><a href=# onClick="window.open('../twt/hit/DIAOYU.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖飞行</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/hunyin/killer.asp','aqjh_win','scrollbars=1,resizable=0,width=640,height=480')">雇佣杀手</A></td><td><a href=# onClick="window.open('../twt/fengxue/fengxue.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">风雪舞剑</a></td></tr>
<tr><td><a href=# onClick="window.open('../twt/findbao/','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">孤岛冒险</A></td><td><a href=# onClick="window.open('../twt/diaoyu/diaoyu.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">休闲钓鱼</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/grade/DIAOYU.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖寻宝</A></td><td><a href=# onClick="window.open('../twt/wokou/index.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">抗击倭寇</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/loot/loot.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">抢劫银行</A></td><td><a href=# onClick="window.open('../twt/hunyin/stunt.asp','aqjh_win','scrollbars=1,resizable=1,width=580,height=350')">合体特技</A></td></tr>
<tr><td><a href=# onClick="window.open('../work/tie/tiemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=370')">矿石炼铁</a></td><td><a href=# onClick="window.open('../work/mine/minemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=470')">江湖采矿</A></td></tr>
<tr><td><a href=# onClick="window.open('../work/famu/famumain.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=370')">江湖伐木</A></td><td><a href=# onClick="window.open('../work/ice/icemain.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=430')">江湖采冰</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/wabao/wabao.asp','aqjh_win','scrollbars=1,resizable=0,width=450,height=340')">挖掘宝藏</A></td><td><a href=# onclick="window.open('../hcjs/dzwq/index.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')" >铸造武器</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/peiyao/peiyao.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')">配药系统</A></td><td><a href=# onClick="window.open('../twt/tao/index.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')">江湖掏金</a></td></tr>
<tr><td><A href="../TOP/TOP47.ASP" target=_blank>在线排行</A></td><td><a href="#"  onClick="javascript:chatwin=window.open('../top/top.htm','','scrollbars=no,resizable=no,width=760,height=400,top=0,left=0')">江湖排行</a></td></tr>
<tr><td><a href=# onClick="window.open('../HCJS/bang/meeting.asp','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">武林大会</a></td><td><a href=../jhshow/jl.asp target=_blank><font color="#FF0000">秀装大赛</a></td></tr>
<tr><td><a href=../jhshow/ target=_blank><font color="#FF0000">虚拟形象</a></td><td><a href=# onClick="window.open('../garden/','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')"><font color="#FF0000">花园大赛</a></td></tr>
</table>
<%
case "江湖集市"
%>
<table width=120>
<tr><td><a href="../CHANGAN/xiaowu.asp" target=_blank>爱的小屋</A></td><td><a href=# onClick="window.open('../hcjs/qr/index.asp','aqjh_win','scrollbars=1,resizable=0,width=650,height=450')">情人小筑</A></td></tr><tr><td><a href=# onClick="window.open('../jhjd/jd.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')">江湖酒店</A></td><td><a href=# onClick="window.open('../hcjs/pub/pub.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">龙门客栈</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/jhjs/hua.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=460')"">江湖花店</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/wuqi.asp','aqjh_win','scrollbars=0,resizable=0,width=650,height=550')">武器商店</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/jhjs/anqi.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=400')">暗器商店</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/yaopu.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖药铺</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/jhjs/vehicle.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">交通工具</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/cw.asp','aqjh_win','scrollbars=0,resizable=0,width=520,height=450')">宠物商店</A></td></tr><tr><td><a href=# onClick="window.open('../jhmp/maidou.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')">卖威力豆</a></td><td><a href=# onClick="window.open('../hcjs/yilao.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">蝶舞山庄</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/yhy/','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">超级妓院</A></td><td><a href=# onClick="window.open('../hcjs/yhyb/index.asp','aqjh_win','scrollbars=1,resizable=1,width=780,height=460')">江湖鸭店</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/qr/fkyy.asp','aqjh_win','scrollbars=1,resizable=0')">妇科中心</A></td><td><a href=# onClick="window.open('../yamen/yiyuan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖医院</A></td></tr><tr><td><a href=# onClick="window.open('../garden/hua.htm','aqjh_win','scrollbars=0,resizable=0,width=670,height=400')">江湖花园</a></td><td><a href=# onClick="window.open('../hcjs/jhjs/dan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖当铺</A></td></tr>
<tr><td><a href=# onClick="window.open('../CHANGAN/fangwu.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=450')">房产开发</A></td><td><a href=# onClick="window.open('../hcjs/mj/mujuan.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">江湖募捐</A></td></tr><tr><td><a href=# onClick="window.open('../jhmp/gry.asp','aqjh_win','scrollbars=0,resizable=0,width=350,height=230')">救助孤儿</A></td><td><a href=# onClick="window.open('../gupiao/stock.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">股票市场</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/dk/Daikuan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">高利贷款</A></td><td><a href=# onClick="window.open('../hcjs/wish/wish.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">江湖祈祷</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/qq/read.htm','aqjh_win','scrollbars=1,resizable=0,width=540,height=480')">宣传奖励</A></td><td><a href=# onClick="window.open('../work/changework.asp','aqjh_win','scrollbars=0,resizable=0,width=520,height=380')">职业转换</A></td></tr><tr><td><a href=# onClick="window.open('../twt/hunyin/','aqjh_win','scrollbars=0,resizable=0,width=550,height=450')">江湖月老</A></td><td><a href=# onClick="window.open('../twt/hunyin/taohun.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">江湖逃婚</A></td></tr>
<tr><td><a href=# onClick="window.open('../bet/betindex.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">江湖赌馆</A></td><td><a href=# onClick="window.open('../hcjs/mc/','aqjh_win','scrollbars=1,resizable=1,width=670,height=400')">清理粪库</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/xuewu/xuetang.htm','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">武馆习武</A></td><td><a href=# onClick="window.open('../twt/dg/dg.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=400')">江湖打工</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/zysc/zysc.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')"><font color="#FF0000">自由市场</a></td><td><a href=# onClick="window.open('../xcj/zh/','aqjh_win','scrollbars=1,resizable=1,width=500,height=300')"><font color="#FF0000">转换市场</a></td></tr>
</table>
<%
case "门派事宜"
%>
<table width=120>
<tr align=center><td><a href=# onClick="window.open('../jhmp/index.asp','aqjh_win','scrollbars=1,resizable=0,width=760,height=450')">七门八派</A></td><td><a href=# onClick="window.open('../jhmp/setwg.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">武功设置</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhmp/addmp/chuangp.asp','aqjh_win','scrollbars=1,resizable=0,width=550,height=450')">自立门派</a></td><td><a href=# onClick="window.open('../jhmp/addmp/rangwei.asp','aqjh_win','scrollbars=1,resizable=1,width=580,height=350')">掌门让位</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../hcjs/room/','aqjh_win','scrollbars=1,resizable=0,width=600,height=450')">自创房间</a></td><td><a href=# onclick="window.open('../jhmp/mympjj.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')" target=_self>门派基金</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/mpxs/mpxs.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=450')">门派薪水</a></td><td><a href=# onclick="window.open('../xcj/mmxl/index.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')" target=_self>门派练功</A></td></tr>
</table>
<%
case "官府衙门"
%>
<table width=120>
<tr align=center><td><a href=# onClick="window.open('../yamen/admin.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')">网管名单</A></td><td><a href=# onClick="window.open('../yamen/laofan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">江湖天牢</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/minan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=470')">命案查询</A></td><td><a href=# onClick="window.open('../jl/jlmain.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=450')">江湖大事</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/laofan1.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">睡眠解救</A></td><td><a href=# onClick="window.open('../hcjs/sy/sy.htm','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">击鼓鸣冤</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/laofang2.asp','aqjh_win','scrollbars=0,resizable=0,width=620,height=400')">查看监禁</A></td><td><a href=../bbs/ target=_blank><font color=red>江湖论坛</a></td></tr>
</table>
<%
case "会员专区"
%>
<table width=120>
<tr><td><a href="hy.asp" target=_blank>购买会员</A></td><td><a href=# onClick="window.open('../jhmp/hy.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">会员列表</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/card/card.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">会员卡片</A></td><td><a href=# onClick="window.open('../jhhy/workauto/ice/icemain.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=430')">会员采冰</A></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/hyxuewu/','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')">会员学武</a></td><td><a href=# onClick="window.open('../jhhy/xiang/index.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">会员拜神</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/workauto/famu/famumain.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=370')">会员伐木</A></td><td><a href=# onClick="window.open('../jhhy/workauto/mine/minemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=450')">会员采矿</A></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/hywabao/wabao.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=450')">会员挖宝</a></td><td><a href=# onClick="window.open('../jhhy/hyjb/index.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">会员炼金</a></td></tr>
</table>
<%
case "其他功能"
%>
<table width=120>
<tr align=center><td><a href="face.asp" scrollbars="0" resizable="0" width="740" height="450">江湖头像</A></td><td><a href=# onClick="window.open('../liwu/xrlw.asp','f3')">新人礼物</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/yxs/yxs.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=470')">月薪领取</A></td><td><a href=# onClick="window.open('../liwu/zllw.asp','f3')">周六礼物</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhmp/bxmp.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">门派排行</A></td><td><a href=# onClick="window.open('../liwu/zwlw.asp','f3')">周五礼物</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/xgn/sm.ASP','aqjh_win','scrollbars=0,resizable=0,width=620,height=400')">功能说明</A></td><td><a href=../bbs/ target=_blank><font color=red>江湖论坛</a></td></tr>
<tr align=center><td><a href=# onClick="window.open('../hcjs/pig/zhu.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">养猪致富</A></td><td><a href=# onClick="window.open('../xcj/maoyi/maoyi.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">贸易市场
<tr align=center><td><a href=# onClick="window.open('../liwu/zmlw.asp','f3')">周末礼品</A></td><td><a href=# onClick="window.open('../liwu/zzlw.asp','f3')">站长礼物</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../liwu/ymlw.asp','f3')">月末奖励</A></td><td><a href=# onClick="window.open('../liwu/jrlw.asp','f3')">节日礼物</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../mhmj/mhmj/index.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">梦幻魔界</A></td><td><a href=# onClick="window.open('../hcjs/zstt/yld.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">幽灵盾牌</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhhy/hyzh/zhuan.asp','aqjh_win','scrollbars=1,resizable=0,width=620,height=400')">转换会员</A></td><td><a href=# onClick="window.open('../jhhy/ydxx/DAJIA.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">造化三国</A></td></tr>

</A></td></tr><tr align=center><td><a href=# onClick="window.open('jhjh/index.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">进化系统</A></td>
</table>
<%
end select
%>
</td></tr></tbody></table> 