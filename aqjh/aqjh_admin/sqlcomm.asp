<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
%>
<html>
<head>
<title>sql指令系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>数据库<br>
  <br>
  <b><font color="#FF0000">注意：此处操作错，将影响到数据库整个系统！<br>
  建议不明白的站长不要使用，非常危险！</font></b> </p>
<p align="left">请输入所要执行的指令：<br>
  如：update 用户 set grade=1,身份='弟子' where 身份&lt;&gt;'掌门' and 门派&lt;&gt;'官府'<br>
  是将所有身份不为掌门，门派不为官府的管理等级清成1级,身份设置成为弟子！！ <br>
  如：delete * from w where i=0<br>
  是将物品数据库中的物品为0的记录删除！<br>
  如：delete * from 武功 where 门派='官府'<br>
  是将官府的武功删除！
</p>
<p> 合并门派之前先执行 <br>
  update 用户 set grade=1, 身份='无' where 门派='门派A'身份='长老' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派A'身份='堂主' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派A'身份='护法' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派B'身份='长老' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派B'身份='堂主' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派B'身份='护法' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派A'身份='长老' <br>
  update 用户 set grade=1, 身份='无' where 门派='门派A'身份='堂主' <br>
  如果不执行会多出很多长老堂主护法的 <br>
  再把被合并的门派的掌门权限拿回来就OK了 </p>
</p>
<p align="left"><br>
  <br>
</p>
<%if aqjh_name=Application("aqjh_user") then%>
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
<OPTION value="delete * from 用户 where allvalue<=2 and month(登录)<>month(now())" >[删除月泡点小于2的用户]</OPTION>
<OPTION value="update 用户 set W8='月饼|10000;' where 姓名='姓名'">[增加一种指定物品]</OPTION>
<OPTION value="update 用户 set allvalue=174050 where allvalue>174050">[把等级大于59的降到59级]</OPTION>
<OPTION value="update 用户 set 银两=0,存款=100000000 where 门派<>'官府' and 银两+存款>100000000">[钱超出1亿的降到1亿]</OPTION>
<OPTION value="update 用户 set W3=W3+'无情剑|1;' where 姓名='用户名'">[给某人一把无情剑]</OPTION>
<OPTION value="update 用户 set zw='蚊子|土龍|2002-7-22 21:05:06|75000|21000|14000|4|2002-9-20 4:13:49' where 姓名='用户名'">[给某人一个宠物]</OPTION>
<OPTION value="update 用户 set 状态='正常',事件原因='无' where lastip='ip地址'">[解禁某人IP]</OPTION>
<OPTION value="update 用户 set 存款=存款+150000000 where grade=5">[给掌们钱的SQL]</OPTION>
<OPTION value="update 用户 set allvalue=allvalue+1000">[给所有人1000点积分]</OPTION>
<OPTION value="update 用户 set 门派='门派名A' where 门派='门派名B'">[合并门派]</OPTION>
<OPTION value="update 用户 set 姓名='新名' where 姓名='旧名'">[给某用户改名]</OPTION>
<OPTION value="update 用户 set 性别='男或女' where 姓名='用户名'">[给某人改性别]</OPTION>
<OPTION value="update 用户 set 法力=法力+1000 where 姓名='用户名'">[调整某人法力]</OPTION>
<OPTION value="update 用户 set w1=' ',w2=' ',w3=' ',w4=' ',w5=' ',w6=' ',w7=' ',w8=' ',w9=' ',w10=' ',z1=' ' where 姓名='回首当年'">[清理用户所有物品]</OPTION>
</SELECT><br>
   <br>
    请输入指令： 
    <input type="text" name="sqlstr" size="50" maxlength="280">
    <br>
    <input type="submit" value="执行" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  </form>
 <%else%>
  <div align="center">您非授权站长，此项禁止使用！</div>
 <%end if%>
</body>
</html>