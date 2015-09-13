<html>
<head>
<title>后台管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false  onselectstart=event.returnValue=false><center>
<table height=30><tr><td></td></tr></table>
<table width=500><tr><td>
<font color=red>+++欢迎您使用『快乐江湖』管理系统+++</font><br><br>
<a href=txt.asp><font color=green>『帮助首页』</font></a>　　<a href=txt_admin.asp><font color=green>『后台说明』</font></a>　　<a href=txt_sql.asp><font color=green>『SQL常规使用』</font></a><br><br>
如果站长对sql指令不熟悉的情况下千万别乱操作，下面告诉大家一些sql方法：<br><br>

1、首先得了解数据库中的数据表,下面是常用的数据表(快乐江湖为例):<br>
　b             物品数据表<br>
　l             管理记录数据表<br>
　npc           npc数据表<br>
　p             门派数据表<br>
　r             房间数据表<br>
　用户          用户数据表<br>
2、一般我们都是对数据库中的用户表进行操作的，在用户表里有很多字段，比如：ID、姓名、性别、体力、内力、武功、等级、grade、法力等等，不过这些大家都知道的，也都很容易记得住的，下面介绍一些物品字段名，这些不好记的：<br>
　w1   药品类<br>
　w2   毒药类<br>
　w3   武器装备<br>
　w4   暗器类<br>
　w5   卡片类<br>
　w6   物品类<br>
　w7   鲜花类<br>
　w8   配药类<br>
　w9   魔宝类<br>
　w10  法器类<br>
　z1   头盔<br>
　z2   装饰<br>
　z3   盔甲<br>
　z4   左手<br>
　z5   右手<br>
　z6   双脚<br>
3、SQL指令最常用的就是修改和删除了，下面是语法：<br>
　修改：update 数据表 set 字段名1=字串或表达式1,字段名2=字串或表达式2... where 表达式<br>
　删除：delete * from 数据表 where 表达式<br><br>
　where后面的表达式说明：比如，想对等级100级以上并且是2级会员的人操作，表达式就写为：等级>100 and 会员等级=2  (如果你连not、and、or都搞不懂的话，那你就不要用这些功能了，如果对操作没有限制，直接省略where以及后面的内容即可<br>
　我们一般都是对用户数据进行处理的，因此以下我们只讲些关于用户处理的sql指令<br>
　例1：update 用户 set 银两=银两+100,性别='男' where 姓名='玲儿'<br>
　　　此sql是修改姓名为玲儿的用户，将她的银两增加100同时把性别改成男，如果还想有多项，可以在性别='男'的后面加逗号再跟上表达式就行了，注意这里面的逗号必须是半角符号,还有一点大家可以看到，为什么银两后面没有'的单引号呢？是因为银两是数值型数据，只有字符型日期型和备注型才可以在两边加单引号'的，只能是单引，不能用双引号的，请牢记<br>
　例2：delete * from 用户 where allvalue<5<br>
　　　此sql是删除总积分小于5的所有用户<br>
4、其它还有好几个命令，这里就不介绍了，比如通过sql添加字段名，建立数据表等，这些用处不大，也比较少，不提倡大家使用的<br>
5、有关常用sql指令，大家可以在后台的sql指令里找到的<br><br>
如果您是正版程序，请观注官方站点，及时更新最新补丁程序<a href=http://www.happyjh.com target=_blank>http://www.happyjh.com</a><br><br>
官方江湖<a href=http://www.happyjh.com target=_blank>http://www.happyjh.com</a><br><br>


----------快乐江湖总站 程序//美工:回首当年------------<br><br>
-------------------2015年5月18日--------------------<br><br>
</td></tr></table>