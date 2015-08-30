<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>『快乐江湖』sql指令系统♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.sqlstr.value;
revisedTitle = addTitle;
document.PostTopic.sqlstr.value=revisedTitle;
document.PostTopic.sqlstr.focus();
return; }
</script>

<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin="0">
<p align="center"><%=Application("sjjh_chatroomname")%>数据库<br>
  <br>
  <b><font color="#FF0000">注意：此处操作错，将影响到数据库整个系统！<br>
  建议不明白的站长不要使用，非常危险！</font></b> </p>
<p align="left">请输入所要执行的指令：<br>
  如：update 用户 set grade=1,身份='弟子' where 身份&lt;&gt;'掌门' and 门派&lt;&gt;'官府'<br>
  是将所有身份不为掌门，门派不为官府的管理等级清成1级,身份设置成为弟子！！ <br>
  如：delete * from w where i=0<br>
  是将物品数据库中的物品为0的记录删除！<br>
  如：delete * from 武功 where 门派='官府'<br>
  是将官府的武功删除！<br>
  如：delete * from 用户 where allvalue&lt;=2 and month(登录)&lt;&gt;month(now())<br>
  是将非本月,泡点小于等于2的用户删除！<br>
<font color=#cc0000>给掌们钱的SQL</font> <br>
update 用户 set 存款=存款+150000000 where grade=5 <br>
<font color=#cc0000>给所有人1000点积分</font> <br>
update 用户 set allvalue=allvalue+1000 <br>
<font color=#cc0000>合并门派</font> <br>
update 用户 set 门派='门派名A' where 门派='门派名B'  <br>

就是说， 把门派B 合并到门派 A里  <br>
执行这个之前先执行  <br>
update 用户 set grade=1, 身份='无' where 门派='门派A'身份='长老' <br> 
update 用户 set grade=1, 身份='无' where 门派='门派A'身份='堂主'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派A'身份='护法'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派B'身份='长老'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派B'身份='堂主'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派B'身份='护法'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派A'身份='长老'  <br>
update 用户 set grade=1, 身份='无' where 门派='门派A'身份='堂主'  <br>
如果不执行会多出很多长老堂主护法的 <br>
再把被合并的门派的掌门权限拿回来就OK了 <br>

<font color=#cc0000>让等级高的降级，积分多少自己计算</font> <br>
update 用户 set allvalue=174050 where allvalue>174050   <br>   
所有用户=59级                       执行条件，等级大于59 <br>
单一的用户 <br>
update 用户 set allvalue=174050 where id=几号 <br>

<font color=#cc0000>所有超过1亿钱的人的钱降到1亿</font> <br>
update 用户 set 银两=0,存款=100000000 where 门派<>'官府' and 银两+存款>100000000 <br>
<font color=#cc0000>给某人一把无情剑</font> <br>
update 用户 set W3='无情剑|1;' where 姓名='小李十一' <br>

<font color=#cc0000>给某人一个宠物，里面的数字要注意哦，不然给的…</font> <br>
update 用户 set zw='蚊子|土龍|2002-7-22 21:05:06|75000|21000|14000|4|2002-9-20 4:13:49' where 姓名='小李十一' <br>

<font color=#cc0000>给某人一个小孩，注意其配偶也要给</font> <br>
update 用户 set boy='星月|baby|2003-5-4 18:44:26|235|40830|11380|20|2003-10-19 21:38:32',boysex='images/boy.gif' where 姓名='回首当年' <br>

<font color=#cc0000>监禁ip解开</font> <br>
update 用户  set 状态='正常',事件原因='无' where lastip='他的ip'<br>

<font color=#cc0000>夺宝初始化</font><br>
update 夺宝 set 冠军=0,领取=false,修练次数=0,宣布=false,时间=now() where 宣布=true  <br>

<font color=#cc0000>月积分奖励</font><br>
update 用户 set 会员金卡=会员金卡+30,金币=金币+100,银币=银币+200,银两=银两+100000000,allvalue=allvalue+1000,金=金+500,木=木+500,水=水+500,火=火+500,土=土+500 where mvalue>=?(查看排行榜) <br>

<font color=#cc0000>总积分奖励</font><br>
update 用户 set 会员金卡=会员金卡+50,金币=金币+200,银币=银币+500,银两=银两+500000000,allvalue=allvalue+2000,金=金+1000,木=木+1000,水=水+1000,火=火+1000,土=土+1000 where allvalue>=?(查看排行榜) <br>

  <br>
</p>
<%if sjjh_name=Application("sjjh_user") then%>
<form method="POST" action="sqlcommok.asp" name="PostTopic">
 <div align="center">
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">快速指令</OPTION> 
<OPTION value="update b set o=1000 where a='物品名' and n<>'无'">[修改竞标收入]</OPTION>
<OPTION value="update 用户 set 银两=0,存款=0 where (银两+存款)>1000000000">[清除资产大于10亿用户资产为0]</OPTION>
<OPTION value="update 用户 set grade=1,身份='弟子' where 门派='游侠'">[清除游侠派管理人员]</OPTION>
<OPTION value="delete * from y where b='官府'" >[删除官府武功]</OPTION>
<OPTION value="delete * from l " >[删除所有操作记录]</OPTION>
<OPTION value="delete * from l where d='人命'" >[删除所有江湖人命记录]</OPTION>
<OPTION value="delete * from j " >[删除所有日记内容]</OPTION>
<OPTION value="delete * from h where c='结婚'" >[删除所有月老记录]</OPTION>
<OPTION value="delete * from h where c='离婚'" >[删除所有离婚记录]</OPTION>
<OPTION value="delete * from t where b='张三'" >[删除男张三的合体技]</OPTION>
<OPTION value="delete * from 用户 where allvalue<=0" >[删除存点allvalue为0的用户]</OPTION>
<OPTION value="delete * from 用户 where 会员金卡>500" >[删除会员金卡大于500的用户]</OPTION>
</SELECT><br>
    <br>
    请输入指令： 
    <input type="text" name="sqlstr" size="50" maxlength="250">
    <br>
    <input type="submit" value="执行" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  </form>
 <%else%>
  <div align="center">您非授权站长，此项禁止使用！</div>
 <%end if%>
<p align="center">『快乐江湖』</p>
</body>
</html>