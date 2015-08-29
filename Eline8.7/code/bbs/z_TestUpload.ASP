<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!--#include file="inc/const.asp" -->
<%
REM 开心辞典 上存题目
REM 我来了 修改为final版
REM 原作者不清楚

	dim ayhuptitlebyte,ayhupanswbyte,ars
	dim alb,atitle,akey1,akey2,akey3,akey4,aok,aup

	stats="开心辞典 上载题目"

	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")

if membername="" or (not founduser) then
	Errmsg=Errmsg+"<br>"+"<li>您没有在开心辞典上载题目的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else
	dim CanUpChoose,CanUpFill
	set ars=server.createobject("adodb.recordset")
	sql="select yhuptitlebyte,yhupanswbyte,CanUpChoose,CanUpFill from [config]"
	set ars=aconn.execute(sql)
	ayhuptitlebyte=ars(0)
	ayhupanswbyte=ars(1)
	CanUpChoose=ars(2)
	CanUpFill=ars(3)
	ars.close

%>		
	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>欢迎 <%=membername%> 上传题目</th></tr>
	<tr><td class=tablebody2>
<%	
	call User_Head()
	
	if cint(CanUpChoose)=0 and (request("action")="" or request("action")="addxzok") then
		errmsg=errmsg+"<br><li>开心辞典没有开放上传选择题的权限"
		founderr=true
	elseif cint(CanUpFill)=0 and (request("action")="addtk" or request("action")="addtkok") then
		errmsg=errmsg+"<br><li>开心辞典没有开放上传填空题的权限"
		founderr=true
	else		
		select case request("action")
			case "addtk"
				call addtk()
			case "addtkok"
				call addtkok()
			case "addxzok"
				call addxzok()	
			case else
				call addxz()
		end select
	end if
	if founderr then call test_err()
		
	response.write "</td></tr></table>"
end if

set aconn=nothing
call activeonline()
call footer()
'========================================================
'上存选择题
sub addxz()
%>
	<form  action="Z_TestUpload.ASP?action=addxzok" method="post">
	  <table align=center cellspacing=1 cellpadding="3" class=tableborder1>
	  <tr><th colspan="2" style='font-weight:normal'><b>上载选择题</b>（注意要把问题和答案写清楚，否则不收录）</th></tr>
	  <tr>
		  <td width="30%" class=tablebody2><b>问题类型：</b></td>
		  <td width="70%" class=tablebody2>&nbsp;<select name="clsid">
		  <option value="" selected>::请选择类型::</option>
<%
	set ars=server.createobject("adodb.recordset")
	sql="select lb,id from [testlb]"
	ars.open sql,aconn,1,1
	while not ars.eof
	  response.write "<option value='"&ars(1)&"'"&">"&ars(0)&"</option>"
	  ars.movenext
	wend
	ars.close
%>
		  </select> <font color="#cc0000">**</font></td>
		</tr>
		<tr>
		  <td class=tablebody1 valign="top"><b>我的问题：</b><br>请把问题写详细，否则不采纳你的题目,请不要超过<font color=blue><%=int(ayhuptitlebyte)%></font>个汉字</td>
		  <td class=tablebody1>&nbsp;<textarea rows="7" name="title" cols="70"></textarea> <font color="#cc0000">&nbsp;**
		  </td>
		</tr>
		<tr>
		  <td class=tablebody2><b>选项1：</b><br>不要超过<font color=blue><%=int(ayhupanswbyte)%></font>个汉字</td>
		  <td class=tablebody2>&nbsp;<input size="70" name="ans1"> <font color="#cc0000">&nbsp;**</font></td>
		</tr>
		<tr>
		  <td class=tablebody2><b>选项2：</b><br>不要超过<font color=blue><%=int(ayhupanswbyte)%></font>个汉字</td>
		  <td class=tablebody2>&nbsp;<input size="70" name="ans2"> <font color="#cc0000">&nbsp;** </font></td>
		</tr>
		<tr>
		  <td class=tablebody2><b>选项3：</b><br>不要超过<font color=blue><%=int(ayhupanswbyte)%></font>个汉字</td>
		  <td class=tablebody2>&nbsp;<input size="70" name="ans3"> <font color="#cc0000">&nbsp;** </font></td>
		</tr>
		<tr>
		  <td class=tablebody2><b>选项4：</b><br>不要超过<font color=blue><%=int(ayhupanswbyte)%></font>个汉字</td>
		  <td class=tablebody2>&nbsp;<input size="70" name="ans4"> <font color="#cc0000">&nbsp;**</font></td>
		</tr>
		<tr>
		  <td class=tablebody2><b>正确答案：</b></td>
		  <td class=tablebody2>&nbsp;<select name="clevel">
			  <option value="" selected>请给出正确答案</option>
			  <option value="1">选项1</option>
			  <option value="2">选项2</option>
			  <option value="3">选项3</option>
		　    <option value="4">选项4</option>
		  </select> <font color="#cc0000">** </font></td>
		</tr>
		<tr>
		  <td class=tablebody2><b>限时：</b><br>预设回答本题限制的时间</td>  
		  <td class=tablebody2>&nbsp;<input name="TimeLimit" type=text value=60> 秒</td>
		</tr>
		<tr>
		  <td class=tablebody2><b>奖金：</b><br>预设本题的奖金</td>  
		  <td class=tablebody2>&nbsp;<input name="JiangJin" type=text value=15> 元</td>
		</tr>				
	  </table>
		<tr><td class=tablebody2 align="center">
			<input type="submit" value="提 交" name="b1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="reset" value="重 写" name="b2">
	  	</td></tr>
	</form>
<%
end sub

sub addxzok()
	if request("clsid")="" or ( not isnumeric(request("clsid")="") ) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择题目类型"
	else			
		alb=cint(request("clsid"))
	end if
	if trim(request("title"))="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>题目不能空,请重新录入新题目"
	elseif len(atitle)>ayhuptitlebyte then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>题目长度超过规定的字数"	
	else				
		atitle=trim(replace(request("title"),"'","＇"))
	end if	
	akey1=trim(replace(request("ans1"),"'","＇"))
	akey2=trim(replace(request("ans2"),"'","＇"))
	akey3=trim(replace(request("ans3"),"'","＇"))
	akey4=trim(replace(request("ans4"),"'","＇"))
	if akey1="" or akey2="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>至少要填写两个选择答案"
	elseif akey3="" and akey4<>"" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请连续填写选择项目"		
	elseif len(akey1)>ayhupanswbyte or len(akey2)>ayhupanswbyte or len(akey3)>ayhupanswbyte or len(akey4)>ayhupanswbyte then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>答案长度超过规定的字数，请重新输入"				
	end if
	aok=request("clevel")
	
	if aok="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定正确答案"		
	end if
	if request("TimeLimit")="" or (not isnumeric(request("TimeLimit"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入正确的限时值，必须为数字"
	end if
	if request("JiangJin")="" or (not isnumeric(request("JiangJin"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入正确的奖金，必须为数字"
	end if
					
	if founderr then
		exit sub
	else
		set ars=server.createobject("adodb.recordset")
		sql="select title from [test] where title like'"&atitle&"'"
		ars.open sql,aconn,1,1
	
		if ars.eof then
			ars.close
			sql="select lb,title,key1,key2,key3,key4,ok,username,tx,TimeLimit,JiangJin from [testtemp]"
			ars.open sql,aconn,1,3
	
			ars.addnew
			ars(0)=alb
			ars(1)=atitle
			ars(2)=akey1
			ars(3)=akey2
			ars(4)=akey3
			ars(5)=akey4
			ars(6)=aok
			ars(7)=membername
			ars(8)=0
			ars(9)=request("TimeLimit")
			ars(10)=request("JiangJin")			
			ars.update
			ars.close

			aconn.execute ("update testuser set userup=userup+1 where username='"&membername&"'")

			sucmsg=sucmsg+"<br>"+"<li>已将本题正常录入试题临时库，请等待管理员审核"
			call suc()
		else
			ars.close
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>你录入的题目原题库已存在，请重新录入新题目"			
		end if
	end if
end sub
'=================================================
'上载填空题
sub addtk()
%>
<form action="Z_TestUpload.ASP?action=addtkok" method="post">
  <table align=center cellspacing=1 cellpadding="3" class=tableborder1>
  <tr><th colspan=2><b>上载填空题</b>（注意要把问题和答案写清楚，否则不收录）</th></tr>
  <tr>
      <td width="30%" class=tablebody2><b>问题类型</b></td>
      <td width="70%" class=tablebody2>&nbsp;<select name="clsid">
      <option value="" selected>::请选择类型::</option>
<%
		set ars=server.createobject("adodb.recordset")
		sql="select lb,id from [testlb]"
		ars.open sql,aconn,1,1
		while not ars.eof
			response.write "<option value='"&ars(1)&"'"&">"&ars(0)&"</option>"
			ars.movenext
		wend
		ars.close
%>
      </select> <font color="#cc0000">**</font></td>
    </tr>
    <tr>
      <td class=tablebody1 valign="top"><b>我的问题：</b><br>请把问题写详细，否则不采纳你的题目,请不要超过<font color=blue><%=int(ayhuptitlebyte)%></font>个汉字</td>
      <td class=tablebody1>&nbsp;<textarea rows="9" name="title" cols="70"></textarea> <font color="#cc0000">&nbsp;**</font></td>
    </tr>
    <tr>
      <td class=tablebody2><b>正确答案：</b><br>不要超过<font color=blue><%=int(ayhupanswbyte/2)%></font><font color="#cc0000">个汉字</td>
      <td class=tablebody2>&nbsp;<input size="70" name="ans1"> <font color="#cc0000">&nbsp;**</font></td>
    </tr>
	<tr>
	  <td class=tablebody2><b>限时：</b><br>预设回答本题限制的时间</td>  
	  <td class=tablebody2>&nbsp;<input name="TimeLimit" type=text value=60> 秒</td>
	</tr>
	<tr>
	  <td class=tablebody2><b>奖金：</b><br>预设本题的奖金</td>  
	  <td class=tablebody2>&nbsp;<input name="JiangJin" type=text value=15> 元</td>
	</tr>	
    </table>
	<tr><td class=tablebody2 align="center">
		<input type="submit" value="提 交" name="b1">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="重 写" name="b2">
  	</td></tr>
	</form>
<%
end sub

sub addtkok()
	if request("clsid")="" or ( not isnumeric(request("clsid")="") ) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择题目类型"
	else			
		alb=cint(request("clsid"))
	end if
	if trim(request("title"))="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>题目不能空,请重新录入新题目"
	elseif len(atitle)>ayhuptitlebyte then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>题目长度超过规定的字数，请重新输入"	
	else				
		atitle=trim(replace(request("title"),"'","＇"))
	end if	
	if trim(request("ans1"))="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>答案不能没有啊，请输入答案"	
	elseif len(akey1)>ayhupanswbyte then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>答案长度超过规定的字数，请重新输入"	
	else
		akey1=trim(replace(request("ans1"),"'","＇"))
	end if
	if request("TimeLimit")="" or (not isnumeric(request("TimeLimit"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入正确的限时值，必须为数字"
	end if
	if request("JiangJin")="" or (not isnumeric(request("JiangJin"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入正确的奖金，必须为数字"
	end if	
	if founderr then
		exit sub
	else
		set ars=server.createobject("adodb.recordset")
		sql="select title from [test] where title like'"&atitle&"'"
		ars.open sql,aconn,1,1
	
		if ars.eof then
			ars.close
			sql="select lb,title,key1,username,tx,TimeLimit,JiangJin from [testtemp]"
			ars.open sql,aconn,1,3
		
			ars.addnew
			ars(0)=alb
			ars(1)=atitle
			ars(2)=akey1
			ars(3)=membername
			ars(4)=1
			ars(5)=request("TimeLimit")
			ars(6)=request("JiangJin")
			ars.update
			ars.close
			aconn.execute ("update testuser set userup=userup+1 where username='"&membername&"'")
			sucmsg=sucmsg+"<br>"+"<li>已将本题正常录入试题临时库，请等待管理员审核"
			call suc()
		else
			ars.close
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>你录入的题目原题库已存在，请重新录入新题目"
		end if
	end if
end sub

sub suc()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>开心辞典 成功信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%>开心辞典
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>继续上存</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=z_test.asp>开心辞典</a></td></tr>
</table>
<%
end sub
%>