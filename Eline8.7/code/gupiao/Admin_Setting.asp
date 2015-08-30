<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-股市控制面板♀wWw.happyjh.com♀</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>股市管理中心</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>本页为管理员专用，您没有管理本页的权限！"
	call endinfo(1)
else 
	select case request("action")
		case "SaveSettimg"		
			call SaveSettimg()
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'关闭数据库
'=====================================
sub main()
	dim KaiHu_Setting,PStock_Setting,User_Setting
	set rs=conn.execute("select top 1 PStock_Setting,KaiHu_Setting,Trade_Setting,User_Setting,AI_Setting,Custom_Setting from [GupiaoConfig] order by id") 
	PStock_Setting=split(rs("PStock_Setting"),"|")
	KaiHu_Setting =split(rs("KaiHu_Setting"),"|")
	'Trade_Setting =split(rs("Trade_Setting"),"|")
	User_Setting  =split(rs("User_Setting"),"|")
	'Custom_Setting=split(rs("Custom_Setting"),"||")
	
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
	<tr>
		<td align="center" colspan="2" height="25" background="Images/Header.gif"><a name=top></a><b>股票设置</b></td>
	</tr>
	<tr>
		<td class="forumRow" colspan="2">
<table width=85% align=center>
<tr>
<td width=50% ><a href="#Gupiao_Setting">[基本设置]</a></td>
<td width=50% ><a href="#Gupiao_Setting1">[防刷新机制]</a></td>
</tr>
<tr>
<td width=50% ><a href="#Gupiao_Setting2">[随机事件]</a></td>
<td width=50% ><a href="#Trade_Setting">[交易设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#PStock_Setting">[个股上市设置]</a></td>
<td width=50% ><a href="#AI_Setting">[机器人(AI)设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#KaiHu_Setting">[开户设置]</a></td>
<td width=50% ><a href="#User_Setting">[帐户设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#Custom_Setting">[自定义位置]</a></td>
<td width=50% ><a href="#Rule_Setting">[股市规则设置]</a></td>
</tr>
</table>		
		</td>
	</tr>	
		
	<form method="POST" action="?action=SaveSettimg">
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Gupiao_Setting></a>基本设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>	
	<tr>
		<td class="forumRow" width="40%"><b>股票市场状态</b><br>股市是否开通</td> 
		<td class="forumRow" width="60%">
			<input type=radio name="Gupiao_Setting(0)" value=1 <%if cint(Gupiao_Setting(0))=1 then%>checked<%end if%>>打开&nbsp;
			<input type=radio name="Gupiao_Setting(0)" value=0 <%if cint(Gupiao_Setting(0))=0 then%>checked<%end if%>>关闭&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>关闭说明</b><br>在股市关闭的情况下显示，支持html语法</td> 
		<td class="forumRow" width="60%">
			<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>自动刷新时间(秒)</b><br>股市首页自动刷新的时间间隔</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(1)" value="<%=Gupiao_Setting(1)%>"> 秒
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>每页显示最多记录</b><br>用于股票所有和分页有关的项目</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(2)" value="<%=Gupiao_Setting(2)%>"> 条
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>是否起用定时开关股市</b></td> 
		<td class="forumRow">
			<input type=radio name="Gupiao_Setting(3)" value=1 <%if cint(Gupiao_Setting(3))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="Gupiao_Setting(3)" value=0 <%if cint(Gupiao_Setting(3))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>股市开放时间</b><br>请确认您已经设置起用定时开关功能<br>起止小时用符号“||”分开</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(4)" value="<%=Gupiao_Setting(4)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>股市名称</b></td> 
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="Gupiao_Setting(5)" value="<%=Gupiao_Setting(5)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><a name=Gupiao_Setting1></a><b>防刷新机制</b><br>如选择打开请填写下面的限制刷新时间</td> 
		<td class="forumRow">
			<input type=radio name="Gupiao_Setting(6)" value=1 <%if cint(Gupiao_Setting(6))=1 then%>checked<%end if%>>打开&nbsp;
			<input type=radio name="Gupiao_Setting(6)" value=0 <%if cint(Gupiao_Setting(6))=0 then%>checked<%end if%>>关闭&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>浏览刷新时间间隔</b><br>填写该项目请确认您打开了防刷新机制</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(7)" value="<%=Gupiao_Setting(7)%>"> 秒
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>防刷新功能有效的页面</b><br>请确认您打开了防刷新功能<br>您指定的页面将有防刷新作用，用户在限定的时间内不能重复打开该页面，具有一定减少资源消耗的作用<br>每个页面名请用“|”符号隔开</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="Gupiao_Setting(8)" value="<%=Gupiao_Setting(8)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>删除不活动用户时间</b><br>可设置删除多少分钟内不活动用户<br>单位：分钟，请输入数字</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(9)" value="<%=Gupiao_Setting(9)%>"> 分钟
		</td>
	</tr>
	<tr>
		<td class="forumRow"><a name=Gupiao_Setting2></a><b>随机事件发生概率</b><br>请输入(0～1)之间数值</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(10)" value="<%=Gupiao_Setting(10)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>随机事件股价上涨的几率</b><br>发生随机事件时股价上涨的几率<br>股价下降的几率=1-股价上涨的几率</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(11)" value="<%=Gupiao_Setting(11)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>保留随机事件记录的条数</b><br>设置为0则保留所有随机事件记录</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="Gupiao_Setting(12)" value="<%=Gupiao_Setting(12)%>"> 条
		</td>
	</tr>
		
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Trade_Setting></a>股票交易设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>是否允许交易</b><br>只有本项打开时股民才能进行股票买卖操作</td> 
		<td class="forumRow">
			<input type=radio name="Trade_Setting(0)" value=1 <%if cint(Trade_Setting(0))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="Trade_Setting(0)" value=0 <%if cint(Trade_Setting(0))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>成交量下限</b><br>当次买卖至少交易量<br>不做限制请设置为0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(2)" value="<%=Trade_Setting(2)%>"> 股		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>成交量上限</b><br>当次买卖至多交易量<br>不做限制请设置为0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(3)" value="<%=Trade_Setting(3)%>"> 股		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>交易次数限制</b><br>同一用户每天交易次数上限<br>不做限制请设置为0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(4)" value="<%=Trade_Setting(4)%>"> 次		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>手续费</b><br>每笔交易收取的手续费的百分数</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(5)" value="<%=Trade_Setting(5)%>">  %	
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>买入时差限制</b><br>同一用户买入同一种股票的时间间隔</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(6)" value="<%=Trade_Setting(6)%>"> 分钟
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>卖出时差限制</b><br>同一用户卖出同一种股票的时间间隔</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(7)" value="<%=Trade_Setting(7)%>"> 分钟
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>买卖时差限制</b><br>同一用户买入某股票后再卖出的时间间隔</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(8)" value="<%=Trade_Setting(8)%>"> 小时
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>涨停上限</b><br>股票价格上涨率大于该值时出现涨停状态</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(9)" value="<%=Trade_Setting(9)%>">
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>跌停下限</b><br>股票价格下滑率大于该值时出现跌停状态</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(10)" value="<%=Trade_Setting(10)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>停牌条件</b><br>股票价格低于该值时出现停牌状态</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(11)" value="<%=Trade_Setting(11)%>"> 元
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>买入对股价的影响</b><br>购入股票时，股票价格 上涨、下降、不变的比例<br>上涨+下降<=1,不变的几率=1-上涨-下降</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Trade_Setting(12)" value="<%=Trade_Setting(12)%>">：<input type=text size="5" name="Trade_Setting(13)" value="<%=Trade_Setting(13)%>">：<%=formatnumber(1-Trade_Setting(12)-Trade_Setting(13),2,true)%>
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>卖出对股价的影响</b><br>抛出股票时，股票价格 下降、上涨、不变的比例<br>上涨+下降<=1,不变的几率=1-上涨-下降</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Trade_Setting(14)" value="<%=Trade_Setting(14)%>">：<input type=text size="5" name="Trade_Setting(15)" value="<%=Trade_Setting(15)%>">：<%=formatnumber(1-Trade_Setting(14)-Trade_Setting(15),2,true)%>
		</td>
	</tr>
																			
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=PStock_Setting></a>个股上市设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>是否开启个股上市申请</b></td> 
		<td class="forumRow">
			<input type=radio name="PStock_Setting(0)" value=1 <%if cint(PStock_Setting(0))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="PStock_Setting(0)" value=0 <%if cint(PStock_Setting(0))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少具备资金</b><br>申请者现有的资金必须达到这个值才能申请个股上市</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="PStock_Setting(1)" value="<%=PStock_Setting(1)%>"> 元
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少上市数目</b><br>最少上市股票数目</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="PStock_Setting(5)" value="<%=PStock_Setting(5)%>"> 股
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>上市公司名称最小长度</b><br>填写数字，不能小于1大于50</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(2)" value="<%=PStock_Setting(2)%>"> byte
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>上市公司名称最大长度</b><br>填写数字，不能小于1大于50</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(3)" value="<%=PStock_Setting(3)%>"> byte 
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>公司简介最大长度</b><br>填写数字，不能小于1大于255</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(4)" value="<%=PStock_Setting(4)%>"> byte 
		</td>
	</tr>

	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=AI_Setting></a>机器人(AI)设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>是否开启机器人</b><br>是否开启机器人事件，只有该项打开后下面的设置才生效</td> 
		<td class="forumRow">
			<input type=radio name="AI_Setting(0)" value=1 <%if cint(AI_Setting(0))=1 then%>checked<%end if%>>开启&nbsp;
			<input type=radio name="AI_Setting(0)" value=0 <%if cint(AI_Setting(0))=0 then%>checked<%end if%>>关闭&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>AI操作几率</b><br>机器人发生买卖股票操作的概率<br>0～1之间的小数，值越大几率越大</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="AI_Setting(1)" value="<%=AI_Setting(1)%>">
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=KaiHu_Setting></a>股市开户设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>是否允许开户</b></td> 
		<td class="forumRow">
			<input type=radio name="KaiHu_Setting(0)" value=1 <%if cint(KaiHu_Setting(0))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="KaiHu_Setting(0)" value=0 <%if cint(KaiHu_Setting(0))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少发贴数</b><br>限制必须达到这个数目才能在股市开户<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(1)" value="<%=KaiHu_Setting(1)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少魅力值</b><br>限制必须达到这个数目才能在股市开户<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(2)" value="<%=KaiHu_Setting(2)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少经验值</b><br>限制必须达到这个数目才能在股市开户<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(3)" value="<%=KaiHu_Setting(3)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>最少金钱值</b><br>限制必须达到这个数目才能在股市开户<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(4)" value="<%=KaiHu_Setting(4)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>原始资本</b><br>赠送给开户者的资本</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="KaiHu_Setting(6)" value="<%=KaiHu_Setting(6)%>"> 元
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=User_Setting></a>股市帐户设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>是否允许从股票帐户提款</b></td> 
		<td class="forumRow">
			<input type=radio name="User_Setting(0)" value=1 <%if cint(User_Setting(0))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="User_Setting(0)" value=0 <%if cint(User_Setting(0))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>是否允许存钱入股票帐户</b></td>  
		<td class="forumRow">
			<input type=radio name="User_Setting(1)" value=1 <%if cint(User_Setting(1))=1 then%>checked<%end if%>>是&nbsp;
			<input type=radio name="User_Setting(1)" value=0 <%if cint(User_Setting(1))=0 then%>checked<%end if%>>否&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>取款上限</b><br>股民每次取款最大金额<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="User_Setting(2)" value="<%=User_Setting(2)%>"> 元
		</td>
	</tr>

	<tr>
		<td class="forumRow"><b>存款上限</b><br>股民每次存款最大金额<br>不做限制请设置为0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="User_Setting(3)" value="<%=User_Setting(3)%>"> 元
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Custom_Setting></a>自定义位置设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>自定义标题</b><br>建议显示的内容不要太长<br><font color=red>支持HTML语法</font></td> 
		<td class="forumRow" width="60%">
			<input type=text name="Custom_Setting(0)" size="74" value="<%=Custom_Setting(0)%>">&nbsp;
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>自定义内容</b><br>可以是图片、文字或滚动消息<br>如果是图片请注意图片的大小(196*52)<br>建议显示的内容不要太长<br><font color=red>支持HTML语法</font></td>   
		<td class="forumRow" width="60%">
			<textarea name="Custom_Setting(1)" cols="75" rows="5"><%=Custom_Setting(1)%></textarea>
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Rule_Setting></a>股市规则设置 [<a href="#top">顶部</a>] [<a href="#bottom">底部</a>]</th>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>涨停规则</b><br>当股票出现涨停状态时是否允许买卖操作</td> 
		<td class="forumRow" width="60%">
			<input type=radio name="Trade_Setting(16)" value=0 <%if cint(Trade_Setting(16))=0 then%> checked <%end if%>>能买不能卖&nbsp;
			<input type=radio name="Trade_Setting(16)" value=1 <%if cint(Trade_Setting(16))=1 then%> checked <%end if%>>能卖不能买&nbsp;
			<input type=radio name="Trade_Setting(16)" value=2 <%if cint(Trade_Setting(16))=2 then%> checked <%end if%>>不能买卖&nbsp;
			<input type=radio name="Trade_Setting(16)" value=3 <%if cint(Trade_Setting(16))=3 then%> checked <%end if%>>不限制&nbsp;
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>跌停规则</b><br>当股票出现跌停状态时是否允许买卖操作</td>   
		<td class="forumRow" width="60%">
			<input type=radio name="Trade_Setting(17)" value=0 <%if cint(Trade_Setting(17))=0 then%> checked <%end if%>>能买不能卖&nbsp;
			<input type=radio name="Trade_Setting(17)" value=1 <%if cint(Trade_Setting(17))=1 then%> checked <%end if%>>能卖不能买&nbsp; 
			<input type=radio name="Trade_Setting(17)" value=2 <%if cint(Trade_Setting(17))=2 then%> checked <%end if%>>不能买卖&nbsp;
			<input type=radio name="Trade_Setting(17)" value=3 <%if cint(Trade_Setting(17))=3 then%> checked <%end if%>>不限制&nbsp; 
		</td> 
	</tr> 

	<tr>
		<td class="forumRow" colspan="2" align="center" height="10"></td>  
	</tr>	
	<tr>
		<td class="forumRow" colspan="2" align="center"><a name=bottom></a><input type="submit" name="Submit" value="提 交"></td>  
	</tr>
	<a name=bottom></a>	
	</form>																							
</table>	
<% 
end sub
'-----------------------
sub SaveSettimg()
	dim KaiHu_Setting,PStock_Setting,Gupiao_Setting,Trade_Setting,User_Setting,StopReadme,AI_Setting,Custom_Setting,Rule_Setting
	Gupiao_Setting=request.Form("Gupiao_Setting(0)")&","&request.Form("Gupiao_Setting(1)")&","&request.Form("Gupiao_Setting(2)")&","&request.Form("Gupiao_Setting(3)")&","&request.Form("Gupiao_Setting(4)")&","&request.Form("Gupiao_Setting(5)")&","&request.Form("Gupiao_Setting(6)")&","&request.Form("Gupiao_Setting(7)")&","&request.Form("Gupiao_Setting(8)")&","&request.Form("Gupiao_Setting(9)")&","&request.Form("Gupiao_Setting(10)")&","&request.Form("Gupiao_Setting(11)")&","&request.Form("Gupiao_Setting(12)")
	PStock_Setting=request.Form("PStock_Setting(0)")&"|"&request.Form("PStock_Setting(1)")&"|"&request.Form("PStock_Setting(2)")&"|"&request.Form("PStock_Setting(3)")&"|"&request.Form("PStock_Setting(4)")&"|"&request.Form("PStock_Setting(5)")
	KaiHu_Setting =request.Form("KaiHu_Setting(0)") &"|"&request.Form("KaiHu_Setting(1)") &"|"&request.Form("KaiHu_Setting(2)") &"|"&request.Form("KaiHu_Setting(3)") &"|"&request.Form("KaiHu_Setting(4)") &"|"&request.Form("KaiHu_Setting(5)") &"|"&request.Form("KaiHu_Setting(6)")
	Trade_Setting =request.Form("Trade_Setting(0)") &"||"&request.Form("Trade_Setting(2)")&"|"&request.Form("Trade_Setting(3)")&"|"&request.Form("Trade_Setting(4)")&"|"&request.Form("Trade_Setting(5)")&"|"&request.Form("Trade_Setting(6)")&"|"&request.Form("Trade_Setting(7)")&"|"&request.Form("Trade_Setting(8)")&"|"&request.Form("Trade_Setting(9)")&"|"&request.Form("Trade_Setting(10)")&"|"&request.Form("Trade_Setting(11)")&"|"&request.Form("Trade_Setting(12)") &"|"&request.Form("Trade_Setting(13)")&"|"&request.Form("Trade_Setting(14)")&"|"&request.Form("Trade_Setting(15)")&"|"&request.Form("Trade_Setting(16)")&"|"&request.Form("Trade_Setting(17)")
	User_Setting  =request.Form("User_Setting(0)")&"|"&request.Form("User_Setting(1)")&"|"&request.Form("User_Setting(2)")&"|"&request.Form("User_Setting(3)")
	AI_Setting    =request.Form("AI_Setting(0)")&"|"&request.Form("AI_Setting(1)")
	StopReadme    =iif(request.Form("StopReadme")<>"","StopReadme='"&request.Form("StopReadme")&"'","StopReadme=null")
	Custom_Setting=request.Form("Custom_Setting(0)")&"||"&request.Form("Custom_Setting(1)")
	'Rule_Setting=request.Form("Rule_Setting(0)")&"|"&request.Form("Rule_Setting(1)")
	set rs=conn.execute("update [GupiaoConfig] set Gupiao_Setting='"&Gupiao_Setting&"',PStock_Setting='"&PStock_Setting&"',KaiHu_Setting='"&KaiHu_Setting&"',User_Setting='"&User_Setting&"',Trade_Setting='"&Trade_Setting&"',"&StopReadme&",AI_Setting='"&AI_Setting&"',Custom_Setting='"&Custom_Setting&"'")
	sucmess="<li>股票信息设置成功"
	call endinfo(1)
end sub
%>