<html><head>
<title>宣传奖励申请</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" rel=stylesheet>
<SCRIPT language=JavaScript>
<!--
function FormCheck(){
        if (document.fm1.oicq.value == "")
        {
        alert("请填写QQ号码！");
		document.fm1.oicq.focus();
        return false;
        }
//判断输入数据是否位数字
        var filter=/[.0-9]/;
        if (!filter.test(document.fm1.oicq.value)) { 
                alert("QQ号码只能使用数字！"); 
                document.fm1.oicq.focus();
                document.fm1.oicq.select();
                return false; 
                } 
		return true;
}
//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF">
<form action="sendsave.asp" onSubmit="return FormCheck();" name="fm1" method="post">
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
    <tr bgcolor="#CCCCCC"> 
      <td colspan="2">◇ <a href="qqlist.asp">申请者列表</a> ◇ <a href="qqlist.asp?id=1">获奖者名单</a> ◇ <a href="qqlist.asp?id=2">处罚名单</a> ◇ <a href="qqlist.asp?id=3">尚未处理名单</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if request("id")=2 then%><font color="#FF0000">请正确填写QQ号码！</font><%end if%></td>
    </tr>
    <tr> 
      <td colspan="2" align="center">爱情江湖QQ宣传奖励申请单</td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
        <hr size="1" noshade>
      </td>
    </tr>
    <tr> 
      <td align="right">您的QQ号：</td>
      <td> 
        <input type="text" name="oicq" maxlength="12">
        <font color="#FF0000">*</font> 注：必须是已经修改了资料后的QQ号</td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="提交">
        <input type="reset" name="Submit2" value="重填">
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <hr size="1" noshade>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <p><b><font color="#FF0000">修改方法：</font></b></p>
        <p><span class="qq">1、将您的QQ的<b>个人详细资料</b>的个人主页栏目修改为：<font color="#FF0000"><b><%=Application("aqjh_homepage")%></b></font><br>
          2、将您的<b>个人说明</b>栏目增加:<%=Application("aqjh_chatroomname")%> <font color="#FF0000"><b><%=Application("aqjh_homepage")%></b></font> 
          当然了加一些修饰性的说明就更好了，例如：最公平的江湖,功能吸引人等 <br>
          3、 奖励措施：凡是修改了QQ资料并保持一个月以上的，送江湖银两2000万，内力200万，武功2万。 </span></p>
        <p><span class="qq">经过我确认后，系统会自动给于奖励。 申请后，系统会自动纪录，大家可以互相监督，发现某些网友作弊请及时检举，我会给于相应的处罚。</span></p>
        <p><span class="qq"><font color="#FF0000">注：</font>请真实填写您的QQ号，如果发现有作弊现象则<font color="#FF0033"><b>删除账号</b></font>永不恢复，对于支持江湖发展帮助宣传江湖的网友我们会给予奖励。凡是修改了QQ资料并保持一个月以上的，送江湖银两2000万，内力200万，武功2万。 
          保持2个月以上者可以留言给站长申请其他的奖励，如果您有几个QQ号或者是朋友的得QQ也可以申请[限制在5个以内]，状态可以重复增加，如果银两超出上限可以用其它的状态值代替奖励。 
          </span></p>
        <p><span class="qq"><font color="#FF0000">处罚</font>：对于乱填写者我们采取罚款等措施，严重者删除账号或封锁IP，对于申请奖励后保持不够一个月者我们也会采取相应的处罚措施收回其增加的状态。</span></p>
      </td></tr></table></form></body></html>