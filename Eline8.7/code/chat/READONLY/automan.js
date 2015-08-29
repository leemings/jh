var huida = new Array();
var wen = new Array();
//wen[x]为在说话中可能出现的问题，而huida[x]就是对话的回答，真正作到智能机器人！
wen[0]="是谁";huida[0] = "我是E线江湖聊天室的小精灵，我会告诉你一些江湖聊天室小技巧哟~";
wen[1]="会员";huida[1] = "会员：为了支持江湖发展请办理会员，江湖会员分为等级制会员、泡点制会员。你可以自由选择加入，具体加入方法请看下面的会员说明！";
wen[2]="站长";huida[2] = "江湖站长：这个江湖的站长是一线天，他的网址是：<a href='http://zhzx.jjedu.org/eline' target='_blank'>http://zhzx.jjedu.org/eline</a>一线天的联系QQ:88617427电话：*****";
wen[3]="程序";huida[3] = "<img src='../chat/pic/st.gif'>哈哈：再在那里胡说，我咬你啊";
wen[4]="等级";huida[4] = "江湖等级：等级分为战斗等级、管理等级。战斗等级使用泡点升级，管理级由掌门、管理员册封而来，管理1级正常用户，2级堂主，3级护法，4级长老，5级掌门，6级以上为管理员！";
wen[5]="密码";huida[5] = "密码：江湖上的密码为不可逆算法，如果你的密码丢了，请找站长联系，请他给你找回密码。";
wen[6]="杀人";huida[6] = "杀人：杀人是要依靠武功、内力、武功修行等值进行的！";
wen[7]="房间";huida[7] = "此套7.0版为多房间版本，你可以点右下角的房间进行必要的切换，选择你要进入的房间";
wen[8]="123";huida[8] = "在E线江湖有很多赚钱的办法：泡分每分钟就10万两;还有<a href='../jhmp/money.asp' target='_blank'>〖江湖工资〗</a>呢,你可以打僵尸—不过你的体力一定要超过僵尸的才行哦，否则你就死悄悄了。";
wen[9]="你好";huida[9] = "在E线江湖有很多赚钱的办法：你可以打僵尸—不过你的体力一定要超过僵尸的才行哦，否则你就死悄悄了。";
wen[10]="我没钱了";huida[10] = "在E线江湖有很多赚钱的办法：你每天都可以去<a href='../jhmp/money.asp' target='_blank'>〖江湖工资〗</a>领取相当的工资的！具体办法呢：你去一次就知道了。";
wen[11]="钱";huida[11] = "在E线江湖有很多赚钱的办法：想要钱就多泡分、多拉人，争取会员最快—去<a href='../jhmp/money.asp' target='_blank'>〖江湖工资〗</a>那里看看吧";
//动态广告
var Banner=new Array();
//广告区内容，自定按格式广告可以设置n条
Banner[0]="欢迎加入【E线江湖】，如果您要加入会员请点<a href='../hy.asp' target='_blank'>这里</a>";
Banner[1]="<a href='../sendphoto/photo.asp' target='_blank'><font color=red>欢迎加入【E线江湖】，请去〖侠客照片〗上传您的靓照</font></a>";
Banner[2]="<a href='../jiaoyou/main.asp?action=search' target='_blank'><font color=red>欢迎加入【E线江湖】，请去〖江湖交友〗寻找您的知心爱人</font></a>";
Banner[3]="<a href='../jhjy/' target='_blank'><font color=blue>新增〖*****〗呵呵，男的当然要去，女的吗...呵呵</font></a>";
Banner[4]="欢迎加入【E线江湖】，如果您要加入会员请点<a href='http://zhzx.jjedu.org/bbs/' target='_blank'>这里给我留言</a>";
Banner[5]="欢迎参与紫华论坛，网址：<a href='http://zhzx.jjedu.org/bbs' target='_blank'>http://zhzx.jjedu.org/bbs</a>如果你有什么意见请告诉我们！";