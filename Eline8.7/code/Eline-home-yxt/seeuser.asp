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
<title>用户数据库管理♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.tiaojian.value;
revisedTitle = addTitle;
document.PostTopic.tiaojian.value=revisedTitle;
document.PostTopic.tiaojian.focus();
return; }
function DoTitlel(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.show.value;
revisedTitle = addTitle;
document.PostTopic.show.value=revisedTitle;
document.PostTopic.show.focus();
return; }

</script>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin="0">
<p align="center"><%=Application("sjjh_chatroomname")%>数据库</p>
<p align="left">请输入条件查询用户：<br>
  如：身份='掌门' 查询身份为掌门的用户。<br>
  再如： 银两&gt;10000 and 武功&gt;10000 查看银两大于10000并且武功大于10000的用户。<br>
  在查询中可以使用：&quot;and&quot; &quot;or&quot; &quot;&gt;&quot; &quot;&lt;&quot; &quot;&lt;&gt;&quot; 
  &quot;&gt;=&quot; &quot;&lt;=&quot; &quot;=&quot; 关系量！<br>
  在查询物品时：查询字段的值无效，如输入：拥有者='小小宇'<br>
  <font color="#0000FF">查询字段：此值为单一值,如：内力、体力等。以下为错误：内力 and 体力 或 内力,体力等</font></p>
<div align="center">用户资料修改程序 </div>
<form method="POST" action="seeuserok.asp" name="PostTopic">
  <div align="center"><br>
    请输入查询条件：
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">选择查询</OPTION> 
<OPTION value="会员=True">[查询泡点制会员]</OPTION>
<OPTION value="会员等级<>0">[查询等级制会员]</OPTION>
<OPTION value="门派='官府' or grade>=6">[官府人员]</OPTION>
<OPTION value="(银两+存款)>1000000000">[资产10亿]</OPTION>
<OPTION value="武功>80000">[武功超8万]</OPTION>
<OPTION value="grade>2 and grade<=5" >[门派管理员]</OPTION>
<OPTION value="身份='掌门' or grade=5">[掌门资料]</OPTION>
<OPTION value="等级>50" >[战级大50]</OPTION>
<OPTION value="regip='输入查询的ip地址' or lastip='输入查询的ip地址'" >[ip查询]</OPTION>
<OPTION value="会员金卡>0">[会员费]</OPTION>
<OPTION value="师傅='小小宇'">[师傅查询]</OPTION>
<OPTION value="金币>=100">[金币大于100]</OPTION>
<OPTION value="会员等级=3">[查询3级会员]</OPTION>
<OPTION value="会员等级>0">[查询所有会员]</OPTION>
</SELECT>
    <input type="text" name="tiaojian" size="50" maxlength="250">
    <br>
    输入将查询字段： 
    <input type="text" name="show" size="10" maxlength="12">
<SELECT name=fs onchange=DoTitlel(this.options[this.selectedIndex].value)> 
<OPTION selected value="">查询内容</OPTION> 
<OPTION value="银两">[银两]</OPTION>
<OPTION value="存款">[存款]</OPTION>
<OPTION value="会员结束">[会员结束]</OPTION>
<OPTION value="会员日期">[会员日期]</OPTION>
<OPTION value="lasttime">[最后上线时间]</OPTION>
<OPTION value="介绍人">[介绍人]</OPTION>
<OPTION value="情人">[情人]</OPTION>
</SELECT>

    <br>
    <input type="submit" value="查询" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  </form>
<p align="center">『快乐江湖』</p>
</body>
</html>