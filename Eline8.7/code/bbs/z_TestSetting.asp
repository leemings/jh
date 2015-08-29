<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<meta http-equiv="Content-Language" content="zh-cn">
<%
'-----------------------
'edit by 我来了/绿水青山
'-----------------------
	dim ars

	stats="开心辞典 辞典管理"
			
	if not master then
		errmsg=errmsg+"<br><li>您没有管理开心辞典的权限，请确认你已经登录并且具有相应的权限！"
		stats="开心辞典 错误信息"
		call nav()
		call head_var(0,0,"开心辞典","Z_test.asp")		
		call dvbbs_error()	
	else
		call nav()
		call head_var(0,0,"开心辞典","Z_test.asp")
%>		
		<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
		<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
		<tr><td class=tablebody2>
<%	
		call Admin_Head()
			
		select case Request("action") 
			case "setting"
				call setting()
			case "OtherSetting"
				call OtherSetting()	
			case "editlb"
				call editlb()
			case "dellb"
				call dellb()
			case "addlb"
				call addlb()
			case "move"
				call move()	
			case else
				call main()
		end select
	end if
%>
		<br>
	</td></tr>
	</table>
<%	
	set ars=nothing
	set aconn=nothing	
	call activeonline()
	call footer()
	
'=================================================						
sub main()			 
	set ars=aconn.execute("select * from [config]")
	dim AnswerSetting,OtherSetting
	AnswerSetting=split(ars("AnswerSetting"),"||")
	if ubound(AnswerSetting)<8 then
		redim preserve AnswerSetting(8)
		AnswerSetting(8)="1"
	end if
	OtherSetting=split(ars("OtherSetting"),"||")	
%>
	<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:97%">
		<tr>
	    	<td align="left" class=TopLighNav1> □ <b><a name=top>参数设置</a></b> <a href="#middle">[中间]</a> <a href="#bottom">[底端]</a></td>
		  	<td align="left" class=TopLighNav1> □ <b>题目类别管理</b> <a href="#middle">[中间]</a> <a href="#bottom">[底端]</a></td>
		</tr>
		<tr>
		<td width="50%" valign="top" class=tablebody1>
		<form method="POST" action="Z_TestSetting.asp?action=setting" name=setting>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
			
				<tr><td align="left" class=tablebody2 colspan="2">◇ 每答一题增减：</td></tr>
				<tr>
					<td width="40%" align="right" class=tablebody1>游戏币：</td>
					<td width="60%" align="left" class=tablebody1><font color=<%=Forum_body(8)%>> － </font><input type="text" name="userp" size="10" value=<%=AnswerSetting(0)%>> 个</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>魅力值：</td>
					<td align="left" class=tablebody1><font color=<%=Forum_body(8)%>> ± </font><input type="text" name="userc" size="10" value=<%=AnswerSetting(1)%>> 点</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>经验值：</td>
					<td align="left" class=tablebody1><font color=<%=Forum_body(8)%>> ± </font><input type="text" name="usere" size="10" value=<%=AnswerSetting(2)%>> 点</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">◇ 答题设置：</td></tr>
				<tr>
					<td align="right" class=tablebody1>每次答题间隔时间：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="dttime" size="10" value=<%=AnswerSetting(3)%>> 分钟</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>每次最多答题数：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="ts" size="10" value=<%=AnswerSetting(4)%>> 题</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>显示正确答案：</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="ShowAnswer" value=1 <%if cint(AnswerSetting(5))=1 then%>checked<%end if%>> 是&nbsp;&nbsp;<input type="radio" name="ShowAnswer" value=0 <%if cint(AnswerSetting(5))=0 then%>checked<%end if%>> 否</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>花钱买知识所需金钱：</td> 
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="AnswerMoney" size="10" value=<%=AnswerSetting(6)%>> ￥</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>答题是否限时：</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="TimeLimit" value=1 <%if cint(AnswerSetting(7))=1 then%>checked<%end if%>> 是&nbsp;&nbsp;<input type="radio" name="TimeLimit" value=0 <%if cint(AnswerSetting(7))=0 then%>checked<%end if%>> 否</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>题目是否支持UBB：</td>
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="UBB" value=1 <%if cint(AnswerSetting(8))=1 then%>checked<%end if%>> 是&nbsp;&nbsp;<input type="radio" name="UBB" value=0 <%if cint(AnswerSetting(8))=0 then%>checked<%end if%>> 否</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">◇ <a name=middle>奖励设置：</a></td></tr>				
				<tr>
					<td align="right" class=tablebody1>用户上传题目奖励：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userjl" size="10" value=<%=ars("userjl")%>> 元</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>上传一题奖游戏币：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userupsinew" size="10" value=<%=ars("userupsinew")%>> 枚</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>每发一帖奖游戏币：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userarticle" size="10" value=<%=ars("userarticle")%>> 枚</td>
				</tr>								
				<tr>
					<td align="right" class=tablebody1>完成答题附加奖励：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="fjjl" size="10" value=<%=ars("fjjl")%>> 元</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">◇ 获得附加奖励条件：</td></tr>				
				<tr>
					<td align="right" class=tablebody1>当次最少答题数：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="dts" size="10" value=<%=ars("dts")%>> 题</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>当次答题正确率：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="zql" size="10" value=<%=ars("zql")%>> %</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">◇ 上传设置：</td></tr>				
				<tr>
					<td align="right" class=tablebody1>允许上传选择题：</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="CanUpChoose" value=1 <%if cint(ars("CanUpChoose"))=1 then%>checked<%end if%>> 是&nbsp;&nbsp;<input type="radio" name="CanUpChoose" value=0 <%if cint(ars("CanUpChoose"))=0 then%>checked<%end if%>> 否</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>允许上传填空题：</td>
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="CanUpFill" value=1 <%if cint(ars("CanUpFill"))=1 then%>checked<%end if%>> 是&nbsp;&nbsp;<input type="radio" name="CanUpFill" value=0 <%if cint(ars("CanUpFill"))=0 then%>checked<%end if%>> 否</td>
				</tr>								
				<tr>
					<td align="right" class=tablebody1>题目最大字数：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="yhuptitlebyte" size="10" value=<%=ars("yhuptitlebyte")%>> Byte</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>答案最大字数：</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="yhupanswbyte" size="10" value=<%=ars("yhupanswbyte")%>> Byte</td>
				</tr>								
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="center" class=tablebody2><input type="submit" value="提交" name="submit">&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重置" name="reset"></td></tr>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=tablebody2>
				友情提示：
				<br>&nbsp;□ <font color="#FF0000">注意：以上设置的参数必须为数字，而且不能为负数哦</font>
				<br>&nbsp;□ 框内为当前执行的参数
				<br>&nbsp;□ 点击"重置"恢复当前执行的参数
				<br>&nbsp;□ 答题间隔时间设为零时，则为答题不限制间隔时间				
				</td></tr>
			</table>
			</form>					
		</td>
		
		<td valign="top" class=tablebody2 width="50%">
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr>
					<td align="center" class=TopLighNav1 width="40">序号</td>
					<td align="center" class=TopLighNav1 width=*>类别名</td>
					<td align="center" class=TopLighNav1 width="100">操作</td>
				</tr>		
<% 
				sql="select lb,id from [testlb]"
				set rs=aconn.execute(sql)
				if rs.eof then
					response.write "<tr><td class=tablebody1 colspan=3>暂时还没有题库类别</td></tr>"
				else
					do while not rs.eof
						response.write "<tr><td align=center class=tablebody1><font color='#FF0000'>"&rs(1)&"</font></td><td align=center class=tablebody1>"&rs(0)&"</td>"
						response.write "<td align=center class=tablebody1><a href='Z_TestSetting.asp?action=dellb&id="&rs(1)&"'"&" onclick=""javascript:{if(confirm('删除类别操作将连同提库内此类别的题同时删除，您确定执行删除操作吗?')){return true;}return false;}""><font color='#FF0000'>删除</a></font></td></tr>"
						rs.movenext
					loop
					rs.close
					set rs=nothing
				end if	
%>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=editlb" name=form3>
				<tr><td align="left" class=TopLighNav1>◇ 修改类别名</td></tr>
				<tr><td align="left" class=tablebody1>&nbsp;序号：<input type="text" name="id" size="3">&nbsp;&nbsp;&nbsp;类别名称：<input type="text" name="LbName" size="10">&nbsp;&nbsp;&nbsp;&nbsp;	<input type="submit" value="修改" name="B3"></td></tr>
				</form>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=addlb" name=form4>
				<tr><td align="left" class=TopLighNav1>◇ 新增类别</td></tr> 
				<tr><td align="left" class=tablebody1>&nbsp;类别名称：<input type="text" name="NewLbName" size="10">&nbsp;&nbsp;&nbsp;<input type="submit" value="添加" name="B4"></td></tr>
				</form>
			</table>

			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=move" name=form5>
				<tr><td align="left" class=TopLighNav1>◇ 题目批量转移</td></tr>
				<tr><td align="left" class=tablebody1>&nbsp;源类别序号：<input type="text" name="OutID" size="3">&nbsp;&nbsp;&nbsp;目标类别序号：<input type="text" name="InID" size="3">&nbsp;&nbsp;&nbsp;&nbsp;	<input type="submit" value="转移" name="B4"></td></tr>
				</form>
			</table>
									  
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=tablebody2>
				友情提示：
				<br>&nbsp;□ <font color="#FF0000">删除类别</font>操作将连同提库内此类别的题同时删除，请慎重
				<br>&nbsp;□ <font color="#FF0000">修改类别</font>名称可直接在上面框内输入与序号相对应的类别名称后点击<font color="#FF0000">修改</font>即可  
				<br>&nbsp;□ <font color="#FF0000">新增类别</font>在上面的框内输入类别名称点击<font color="#FF0000">添加</font>即可
				<br>&nbsp;□ <font color="#FF0000">批量转移</font>的作用是把属于某个分类的所有题目移到另外一个分类中
				</td></tr>
			</table>
			
			</td>
		</tr>

		<tr>
	    <td align="left" class=TopLighNav1> □ <b><a name=bottom>其他设置</a></b> <a href="#middle">[中间]</a> <a href="#top">[顶端]</a></td>
		<td align="center" class=TopLighNav1><b></b></td>
		</tr>
		<form method="POST" action="Z_TestSetting.asp?action=OtherSetting" name=form2>
		<tr>
		<td width="50%" valign="top" class=tablebody1>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
											
				<tr>
					<td align="left" class=tablebody1 width="45%">辞典首页是否显示图片：</td> 
					<td align="left" class=tablebody1 width="55%">&nbsp;<input type="radio" name="ShowPic" value=1 <%if cint(OtherSetting(0))=1 then%>checked<%end if%>> 图片&nbsp;&nbsp;<input type="radio" name="ShowPic" value=2 <%if cint(OtherSetting(0))=2 then%>checked<%end if%>> 广告&nbsp;&nbsp;<input type="radio" name="ShowPic" value=3 <%if cint(OtherSetting(0))=3 then%>checked<%end if%>> 滚动消息</td>
				</tr>
				<tr>
					<td align="left" class=tablebody1>辞典首页图片路径：</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="text" name="PicPath" value="<%=OtherSetting(1)%>" size="25"></td>
				</tr>				
				<tr>
					<td align="left" class=tablebody1>答题龙虎榜显示记录条数：</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="text" name="PaiXingTop" value="<%=OtherSetting(2)%>"> 条</td>  
				</tr>				
			</table>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="center" class=tablebody2><input type="submit" value="提交" name="submit">&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重置" name="reset"></td></tr>
			</table>			
			
		</td>
		<td width="50%" valign="top" class=tablebody1>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=TopLighNav1>◇ 辞典首页广告代码[<font color=red>支持HTML</font>]</td></tr>							
				<tr>
					<td align="left" class=tablebody1 width="100%">
					<textarea name="kxcd_ad" cols="65" rows="5"><%=ars("ad")%></textarea> 
					</td>
				</tr>
				
			</table>
			<table border=0><tr><td height="3"></td></tr></table>		
		</td>
		</tr>
		</form>		
		</table>
<%	
	ars.close
end sub

sub setting()
	if request("userp")="" or (not isnumeric(request("userp"))) then
		errmsg=errmsg+"<br><li>游戏币不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userp")<0 then 
		errmsg=errmsg+"<br><li>游戏币不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userc")="" or (not isnumeric(request("userc"))) then
		errmsg=errmsg+"<br><li>魅力值不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userc")<0 then 
		errmsg=errmsg+"<br><li>魅力值不能为负数，请重新输入"
		founderr=true	
	end if
	if request("usere")="" or (not isnumeric(request("usere"))) then
		errmsg=errmsg+"<br><li>经验值不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("usere")<0 then 
		errmsg=errmsg+"<br><li>经验值不能为负数，请重新输入"
		founderr=true	
	end if	
	if request("dttime")="" or (not isnumeric(request("dttime"))) then
		errmsg=errmsg+"<br><li>答题间隔时间不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("dttime")<0 then 
		errmsg=errmsg+"<br><li>答题间隔时间不能为负数，请重新输入"
		founderr=true	
	end if
	if request("ts")="" or (not isnumeric(request("ts"))) then
		errmsg=errmsg+"<br><li>每次最多答题数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("ts")<0 then 
		errmsg=errmsg+"<br><li>每次最多答题数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("AnswerMoney")="" or (not isnumeric(request("AnswerMoney"))) then
		errmsg=errmsg+"<br><li>花钱买知识所需金钱不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("AnswerMoney")<0 then 
		errmsg=errmsg+"<br><li>花钱买知识所需金钱不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userjl")="" or (not isnumeric(request("userjl"))) then
		errmsg=errmsg+"<br><li>用户上传题目奖励不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userjl")<0 then 
		errmsg=errmsg+"<br><li>用户上传题目奖励不能为负数，请重新输入"
		founderr=true	
	end if
	if request("fjjl")="" or (not isnumeric(request("fjjl"))) then
		errmsg=errmsg+"<br><li>完成答题附加奖励不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("fjjl")<0 then 
		errmsg=errmsg+"<br><li>完成答题附加奖励不能为负数，请重新输入"
		founderr=true	
	end if
	if request("zql")="" or (not isnumeric(request("fjjl"))) then
		errmsg=errmsg+"<br><li>当次答题正确率不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("zql")<0 or request("zql")>100 then 
		errmsg=errmsg+"<br><li>当次答题正确率不能大于100或者小于0，请重新输入"
		founderr=true	
	end if
	if request("dts")="" or (not isnumeric(request("dts"))) then
		errmsg=errmsg+"<br><li>当次最少答题数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("dts")<0 then 
		errmsg=errmsg+"<br><li>当次最少答题数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userupsinew")="" or (not isnumeric(request("userupsinew"))) then
		errmsg=errmsg+"<br><li>每发一帖奖游戏币不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userupsinew")<0 then 
		errmsg=errmsg+"<br><li>每发一帖奖游戏币不能为负数，请重新输入"
		founderr=true	
	end if
	if request("userarticle")="" or (not isnumeric(request("userarticle"))) then
		errmsg=errmsg+"<br><li>上传一题奖游戏币不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("userarticle")<0 then 
		errmsg=errmsg+"<br><li>上传一题奖游戏币不能为负数，请重新输入"
		founderr=true	
	end if
	if request("yhuptitlebyte")="" or (not isnumeric(request("yhuptitlebyte"))) then
		errmsg=errmsg+"<br><li>题目最大字数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("yhuptitlebyte")<0 then 
		errmsg=errmsg+"<br><li>题目最大字数不能为负数，请重新输入"
		founderr=true	
	end if
	if request("yhupanswbyte")="" or (not isnumeric(request("yhupanswbyte"))) then
		errmsg=errmsg+"<br><li>答案最大字数不能为空或者非数字，请重新输入"
		founderr=true
	elseif request("yhupanswbyte")<0 or request("yhupanswbyte")>200then 
		errmsg=errmsg+"<br><li>答案最大字数不能小于0或者大于200，请重新输入"
		founderr=true	
	end if										
	
	if founderr then
		call test_err()
		exit sub
	end if
	
	if cint(request("dts"))>cint(request("ts")) then
		errmsg=errmsg+"<br><li>哇，晕了，当次最少答题数 比 本次最多答题数 还要多啊，谁能拿到奖励哦？"
		call test_err()
		exit sub
	end if		
	
	dim AnswerSetting 
	AnswerSetting=request("userp")&"||"&request("userc")&"||"&request("usere")&"||"&request("dttime")&"||"&request("ts")&"||"&request("ShowAnswer")&"||"&request("AnswerMoney")&"||"&request("TimeLimit")&"||"&request("UBB")
	aconn.execute("update [config] set AnswerSetting='"&AnswerSetting&"',userjl="&request("userjl")&",fjjl="&request("fjjl")&",zql="&request("zql")&",dts="&request("dts")&",userupsinew="&request("userupsinew")&",userarticle="&request("userarticle")&",yhuptitlebyte="&request("yhuptitlebyte")&",yhupanswbyte="&request("yhupanswbyte")&",CanUpChoose="&request("CanUpChoose")&",CanUpFill="&request("CanUpFill")&"")
	sucmsg=sucmsg+"<br><li>参数设置成功，请返回进行其他操作"
	call suc()
end sub

sub OtherSetting()
	if request("ShowPic")=1 and request("PicPath")="" then
		errmsg=errmsg+"<br><li>请输入图片路径"
		founderr=true
	end if
	if request("PaiXingTop")="" then
		errmsg=errmsg+"<br><li>请输入答题龙虎榜显示记录条数"
		founderr=true	
	elseif not isnumeric(request("PaiXingTop")) then
		errmsg=errmsg+"<br><li>答题龙虎榜显示记录条数必须输入数字"
		founderr=true
	elseif request("PaiXingTop")<=0 then
		errmsg=errmsg+"<br><li>答题龙虎榜显示记录条数必须大于零"
		founderr=true
	end if	
	if founderr then
		call test_err()
		exit sub
	end if	
			
	dim OtherSetting  
	OtherSetting=request("ShowPic")&"||"&request("PicPath")&"||"&request("PaiXingTop")
	aconn.execute("update [config] set OtherSetting='"&OtherSetting&"',ad='"&CheckStr(request("kxcd_ad"))&"'")
	sucmsg=sucmsg+"<br><li>参数设置成功，请返回进行其他操作"
	call suc()
end sub

sub editlb()
	if request("id")="" or isnull(request("id")) then
		errmsg=errmsg+"<br><li>请输入要修改的类别的序号"
		founderr=true
	elseif not isnumeric(request("id")) then
		errmsg=errmsg+"<br><li>序号必须是数字"	
		founderr=true
	end if	
	if request("LbName")="" then
		errmsg=errmsg+"<br><li>请输入新类别名称"
		founderr=true
	end if 
	if founderr then
		call test_err()
		exit sub
	end if
    set ars=aconn.execute("select lb from [testlb] where id="&request("id"))
	if ars.eof and ars.bof then
		ars.close
		errmsg=errmsg+"<br><li>没有相关类别，请输入正确的类别序号"
		founderr=true
		call test_err()
		exit sub		
	elseif trim(ars(0))=trim(request("LbName")) then
		ars.close
		errmsg=errmsg+"<br><li>你输入的类别名称已经存在，请重新输入"
		founderr=true
		call test_err()
		exit sub						
	end if
	ars.close
			
	aconn.execute("update [testlb] set lb='"&trim(replace(request("LbName"),"'",""))&"' where id="&request("id"))
	
	sucmsg=sucmsg+"<br><li>类别名修改成功，请返回进行其他操作"
	call suc()	
end sub

sub dellb()
	aconn.execute ("delete from [testlb] where id="&request("id"))
	aconn.execute ("delete from [test] where lb="&request("id"))
	sucmsg=sucmsg+"<br><li>类别删除成功，属于该类别的题目也已经删除，请返回进行其他操作"
	call suc()
end sub

sub addlb()
	if request("NewLbName")="" then
		errmsg=errmsg+"<br><li>请输入类别名称"
		call test_err()
		exit sub
	end if 
	
	set ars=aconn.execute("select lb from [testlb] where lb='"&trim(replace(request("NewLbName"),"'",""))&"'")
	if not (ars.eof and ars.bof) then
		errmsg=errmsg+"<br><li>你输入的类别名称已经存在，请重新输入"
		founderr=true
		ars.close
		call test_err()
		exit sub						
	end if
	ars.close		
	aconn.execute("insert into [testlb](lb) values('"&trim(replace(request("NewLbName"),"'",""))&"')")
	
	sucmsg=sucmsg+"<br><li>成功增加新类别，请返回进行其他操作"
	call suc()
end sub

sub move()
	if request("OutID")="" then
		errmsg=errmsg+"<br><li>请输入移出的分类的序号"
		founderr=true
	elseif not isInteger(request("OutID")) then
		errmsg=errmsg+"<br><li>移出的分类的序号不正确，请重新输入"
		founderr=true		
	end if 
	if request("InID")="" then
		errmsg=errmsg+"<br><li>请输入移入的分类的序号"
		founderr=true
	elseif not isInteger(request("InID")) then
		errmsg=errmsg+"<br><li>移入的分类的序号不正确，请重新输入"
		founderr=true		
	end if 
		
	if founderr then
		call test_err()
		exit sub
	end if
		
	set ars=aconn.execute("select id from [testlb] where id="&request("InID"))
	if ars.eof and ars.bof then
		errmsg=errmsg+"<br><li>目标类别不存在"
		founderr=true
	end if
	ars.close
	
	set ars=aconn.execute("select id from [test] where lb="&request("OutID"))
	if ars.eof and ars.bof then
		errmsg=errmsg+"<br><li>在源类别中没有任何题目或者源类别并不存在"
		founderr=true
	end if
	ars.close	
	
	if founderr then 
		call test_err()
	else	
		aconn.execute("update [test] set lb="&request("InID")&" where lb="&request("OutID"))
		sucmsg=sucmsg+"<br><li>操作成功，请返回进行其他操作"
		call suc()
	end if
end sub
%>