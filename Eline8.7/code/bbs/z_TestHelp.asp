<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<%
	dim ars,auserjl,azql,adts,afjjl,auserupsinew,auserarticle
	dim AnswerSetting,OtherSetting
	stats="开心辞典 知识问答帮助"
	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")
	call activeonline()
	
	sql="select AnswerSetting,OtherSetting,userjl,zql,dts,fjjl,userarticle,userupsinew from [config]"
	set ars=aconn.execute(sql)
	AnswerSetting=split(ars(0),"||")	'0每回答一题所需游戏币个数 ||1每回答一题魅力值增减值||2每回答一题经验值增减值||3两次玩之间的时间间隔||4答题数||5是否显示正确答案||6花钱买知识所需金钱||7答题是否限时||8题目是否支持UBB
	if ubound(AnswerSetting)<8 then
		redim preserve AnswerSetting(8)
		AnswerSetting(8)="1"
	end if
	OtherSetting=split(ars(1),"||")		'0是否显示图片 || 1图片路径 || 2答题龙虎榜显示记录条数
	auserjl=ars(2)			'上传奖励金额 
	azql=ars(3)				'获得附加奖励条件 当次答题正确率
	adts=ars(4)				'获得附加奖励条件 当次最少答题数
	afjjl=ars(5)			'附加奖励金额
	auserarticle=ars(6)		'每发一帖奖游戏币
	auserupsinew=ars(7)		'上传一题奖游戏币   
	ars.close
	set ars=nothing
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>开心辞典---知识问答帮助</th></tr>
<tr><td class=tablebody2>

  <table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
    <tr>
          <td width="100%" class=TopLighNav1>
          <p align="center"><b><font color="#FF0000">知识问答帮助</font></b></td>
    </tr>
    <tr>
          <td width="100%" class=tablebody2>
		  <br>
		  <p align="center"><font color=blue>开心辞典 V1.2 版</font></p>
		  &nbsp;<span lang="zh-cn">1、如何增加游戏币：</span><p>
          <span lang="zh-cn">　　A：发帖挣游戏币：每发一帖，可挣游戏币 <font color=red><b><%=auserarticle%></b></font> 枚</span>,<span lang="zh-cn">自动更新。</span></p>
          <p><span lang="zh-cn">　　B：上传题目挣游戏币：每上传一道题目并通过审核，可挣游戏币 <font color=red><b><%=auserupsinew%></b></font> 枚。</span></p>
          <p>&nbsp;<span lang="zh-cn">2、知识问答有关设置参数：</span></p>
          <p><span lang="zh-cn">　　游戏币：&nbsp;每答一题<b>&nbsp;<font color="#FF0000">－<%=AnswerSetting(0)%></font></b></span></p> 
          <p><span lang="zh-cn">　　经验值：</span><span lang="en-us">&nbsp;</span>对<span lang="en-us"><font color="#FF0000"><b> +</b></font></span><b><font color="#FF0000"><%=AnswerSetting(2)%></font></b> 错<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(2)%></font></b></p>
          <p><span lang="zh-cn">　　魅力值：</span><span lang="en-us">&nbsp;</span>对<b><font color="#FF0000"><span lang="en-us"> +</span><%=AnswerSetting(1)%></font></b> 错<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(1)%></font></b></p>
          <p><span lang="zh-cn">　　现　金：</span><font color=red>按题目不同相应增减</font></p>
          <p><span lang="zh-cn">　　每次答题数：</span><font color=red><b><%=AnswerSetting(4)%></b></font></p>
          <p><span lang="zh-cn">　　每次答题时间间隔：</span><font color=red><b><%=AnswerSetting(3)%></b></font></p>
          <p><span lang="zh-cn">　　得到附加奖励<font color=red><b><%=afjjl%></b></font>元的条件：答对数不少于：<font color=red><b><%=adts%></b></font>  正确率不低于：</span><font color=red><b><%=azql%></b>%</font></p>
          <p><span lang="zh-cn">　　显示正确答案：</span><font color=red><%if cint(AnswerSetting(5))=1 then%>显示<%else%>花钱买知识<%end if%></font></p> 
		  <p><span lang="zh-cn">　　看答案所需金钱：</span><font color=red><%=AnswerSetting(6)%> 元</font></p>
		  <p><span lang="zh-cn">　　答题时间限制：</span><font color=red><%if cint(AnswerSetting(7))=1 then%>打开<%else%>关闭<%end if%></font></p>
		  <p><span lang="zh-cn">　　题目是否支持UBB标签：</span><font color=red><%if cint(AnswerSetting(8))=1 then%>支持<%else%>不支持<%end if%></font></p>
		  <p><span lang="zh-cn">　　出题顺序：</span><font color=red>随机顺序</font></p>
		  &nbsp;<span lang="zh-cn">3、如何开始游戏：</span><p> 
		  <p><span lang="zh-cn">　　点击题目类型即可进入知识问答比赛</span></p> 
		  <p><span lang="zh-cn">　　</span></p>     
		  </td>
    </tr> 
    <tr>
    <td width="100%" align="center" class=tabletitle2><a href=javascript:history.go(-1)>返回上一页</a></td> 
    </tr>
</table>
</td></tr></table>
<%
set aconn=nothing
call footer()
%>