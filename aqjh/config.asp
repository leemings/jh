<%
Application("aqjh_chatbgcolor")="4a79b5"      '聊天室背景颜色代码
Application("aqjh_chatimage")="1.jpg"      '聊天室外部边框图片
Application("aqjh_chatcolor")="effaff"      '聊天室文字区背景色
Application("aqjh_iplocktime")=1200000      '锁IP时间
Application("aqjh_maxtimeout")=20      '泡点超时时间，不要设置太长
Application("aqjh_maxnpc")=10      'npc总在线人数
Application("aqjh_npcwp")=500      'npc暴涨比较值
Application("aqjh_npcff")=10      'npc发放需要的等级（管理级）
Application("aqjh_baowuyin")=200000      '每一级修练所得到的银两
Application("aqjh_baowuxl")=25      '每次修练的级数，如100级的需修练4次！他所得到的上限就是4！这个值是用战斗等级取整的
Application("aqjh_killman")=6      '江湖一天的杀人数，夺宝一次算一次杀人
aqjh_wgsx=2500      '每升一级的武功加值,此值当aqjh_sx=1时才有作用
aqjh_nlsx=1200      '每升一级的内力加值,此值当aqjh_sx=1时才有作用
aqjh_tlsx=3000      '每升一级的体力加值,此值当aqjh_sx=1时才有作用
aqjh_gjsx=1600      '每升一级的攻击加值,此值当aqjh_sx=1时才有作用
aqjh_fysx=2000      '每升一级的防御加值,此值当aqjh_sx=1时才有作用
Application("aqjh_chat_maxpeople")=500      '江湖最大人数限制
Application("aqjh_closedoor")=0      '聊天室关门，1为关门
Application("aqjh_baowu")=1      '是否有宝物当为0时没有宝物！宝物出现是会影响聊天室速度的，服务器不好关闭！
aqjh_sx=1      '是否有上限0没有上限，1有上限
Application("aqjh_automanname")="爱情天使"      '聊神名称设置
Application("aqjh_userout")="“%%向大家挥挥手说声再见，然后离开了聊天室！”"      '退出时聊天室显示的信息
aqjh_dieip="218.75.59.42;218.62.122.70;219.134.19.203;61.141.144.113;218.17.61.239;218.17.59.90;"      'ip永久封锁使用前按格式书，多个用分号隔开，后面请保留有分号
Application("aqjh_baowuname")="爱情水晶石"      '宝物名及宝物说明
Application("aqjh_baowusm")="这宝物是霄逸寻便大江南北，才辛辛苦苦找到的。但是想不道，不知道让司空摘星盗走，唉。。。"      '宝物在聊天室中显示信息
'***************************************************
'*以下江湖管理设置，切记不要改错，否则无法正常运行 *
'***************************************************
Application("aqjh_chatroomname")="爱情江湖"      '江湖名称
Application("aqjh_homepage")="http://www.7758530.com"      '江湖主页
Application("aqjh_sn")="love888"      '江湖序列号，正版标志，修改后后果自负
Application("aqjh_user")="千万"      '江湖站长
Application("aqjh_qq")="51726805"      '站长qq
Application("aqjh_email")="119yejin@163.com"      '站长email
Application("aqjh_admin")="永不放弃"      '设置10级聊管，多个用逗号隔开
Application("hidden_admin")="永不放弃"      '隐身人员，多个用逗号隔开
Application("aqjh_slbox")="永不放弃"      '可以查看私聊的人，多个用逗号隔开
Application("aqjh_admin_send")="||"      '财神爷，多个用|和|隔开
Application("aqjh_guibin")="||"      '江湖贵宾(贵宾不参与任何江湖恩怨)，多个用|和|隔开
Application("aqjh_adminuser")="永不放弃"      '站长管理用户名
Application("aqjh_adminkey")="123456"      '站长管理密码
Application("tuziji_DoReflashPage") = False  '页面防止刷新'
'*********************************************************
'*如果要设置广告内容请见：chat/jhchat.asp文件开始处
'*泡分的点数可以由自己设置，详见：paodian.asp文件开始处
'修改招式等级修改chat/sjfunc/czdj.asp
%>
