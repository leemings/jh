<!--#include file="conn.asp"-->
<!--#include file="z_testconn.asp"-->

<!-- #include file="inc/const.asp" -->
<%
'------------------------
'script writed by 我来了
'2003-01-06     幻海晴空
'------------------------
	stats="开心辞典 用户管理"
	
	if (not master) and Request("action")<>"" then
		errmsg=errmsg+"<br><li>您没有在开心辞典管理用户的权限，请确认你已经登录并且具有相应的权限！"
		founderr=true
		stats="开心辞典 错误信息"
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")		
		call dvbbs_error()
	else
		call nav()
		call head_var(0,0,"开心辞典","z_test.asp")
		dim ars
		select case request("action")
			case "LayIn"
				call LayIn()
			case else		
				call main()
		end select
	end if
	set aconn=nothing
	call activeonline()
	call footer()
'================================================
'-----------------开心辞典用户列表--------------
sub main()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>欢迎 <%=membername%> 进入开心辞典 管理中心</th></tr>
	<tr><td class=tablebody2> 
    <%call Admin_Head()%>
	
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%"> 
		<tr>
		  <th colspan="3" align=left>题库批量导入功能</th>
		</tr>
		<tr>
			<td colspan="3" align="center" class=tablebody1>注意:请认真按要求填写以下内容!</td>
		</tr>
		<form action="?action=LayIn" method=post>
		<tr>
			<td width="20%" height="1" align="right" valign="middle" class=tablebody1>源数据库名：</td>
			<td width="60%" height="1" class=tablebody1><input type="text" name="DBname" size="30"> 如：testchen.mdb　或　plus/testchen.mdb</td>
			<td width="20%" height="3" rowspan="4" class=tablebody1>
			&nbsp;<font color=blue>□</font> 导入题库前，请查看新题库数据库的名称与记录题目和类别的表名，并填写到左边表单中<br>
			&nbsp;<font color=red>□</font> 本操作比较占用服务器资源，建议在本地进行，或者选择服务器人少的时候进行 
			</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">源类别表名：</td> 
			<td class=tablebody1><input type="text" name="LBTable" size="30" value="TestLb"> 如：TestLb</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">源题目表名：</td>
			<td class=tablebody1><input type="text" name="TiMuTable" size="30" value="Test"> 如：Test</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">导入选项：</td>
			<td class=tablebody1><input type="radio" name="Setting"  value="0" checked>在原题库的基础上追加 &nbsp;<input type="radio" name="Setting"  value="1">删除旧题目后导入</td>
		</tr>				
		<tr>
			<td class=tablebody1 align="right" valign="middle"></td>
			<td class=tablebody1 colspan="2"><input type="submit" name="Submit" value="导入"></td>
		</tr>
		</form>
	</table>
	<br>
	</td>
	</tr>
</table>
<% 
end sub

sub LayIn()
	dim DBname,LBTable,TiMuTable,Tconn
	DBname=replace(trim(request.Form("DBname")),"'","")
	LBTable=replace(trim(request.Form("LBTable")),"'","")
	TiMuTable=replace(trim(request.Form("TiMuTable")),"'","")
	
	if DBname="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入的数据库名"
		call test_err()
		exit sub
	end if
	if LBTable="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入的类别表名"
		call test_err()
		exit sub
	end if
	if TiMuTable="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入的题目表名"
		call test_err()
		exit sub
	end if

	'取出新的数据进行更新
	on error resume next	
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBname)		'连接导入的数据库
	
	if err.number=0 then
		dim trs,nrs,NewLbID,OldTitleID,OldLbID
		set nrs=tconn.execute("select lb,id from "&LBTable&" order by id")         '打开导入的类别表
		
		if cint(request.Form("Setting"))=1 then
			'取出原题库中题目id和类别id
			OldTitleID=aconn.execute("select top 1 id from [test] order by id desc")(0)
			OldLbID=aconn.execute("select top 1 id from [testlb] order by id desc")(0)
		end if  
		if err.number=0 then
		
			do while not nrs.eof 
				aconn.execute("insert into [testlb] (lb) values('"&nrs(0)&"')")
				NewLbID=aconn.execute("select top 1 id from [testlb] order by id desc")(0)
				set trs=tconn.execute("select * from "&TiMuTable&" where lb="&nrs(1))		'打开导入的题目表
				if err.number<>0 then exit do	
				do while not trs.eof
					aconn.execute("insert into [test] (lb,jj,[time],title,key1,key2,key3,key4,ok,tx,tcount,okcount,username) values ("&NewLbID&","&trs("jj")&","&trs("time")&",'"&replace(trs("title"),"'","＇")&"','"&trs("key1")&"','"&trs("key2")&"','"&trs("key3")&"','"&trs("key4")&"','"&trs("ok")&"',"&trs("tx")&","&trs("tcount")&","&trs("okcount")&",'"&trs("username")&"')")
					trs.movenext
				loop
				trs.close
				nrs.movenext
			loop
			nrs.close
			set nrs=nothing		
			set trs=nothing
			
			if cint(request.Form("Setting"))=1 and err.number=0 then
				'删除已有题目和类别
				aconn.execute("delete from [test] where id<="&OldTitleID&"")    
				aconn.execute("delete from [testlb] where id<="&OldLbID&"") 
				sucmsg=sucmsg+"<br><li>旧题库已经删除"
			end if
			
		end if
	end if	  
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li>数据库操作失败，请确认输入无误以及数据库结构是否一致；出错信息如下："
		ErrMsg=ErrMsg+"<Br>"+"<li>"&err.Description
		err.clear
		call test_err()
	else
		dim LBnum,TitleNum,Total,LayInNum
		LBnum=tconn.execute("select count(id) from "&LBTable&"")(0)
		Total=tconn.execute("select count(id) from "&TiMuTable&"")(0)
		LayInNum=tconn.execute("select count(id) from "&TiMuTable&" where lb in (select id from "&LBTable&")")(0)
		
		sucmsg=sucmsg+"<br><li>新题库导入成功"
		sucmsg=sucmsg&"<br><li>导入情况：源题库 共有<font color=red>"&LBnum&"</font>个分类，总题数：<font color=red>"&Total&"</font>题，成功导入：<font color=red>"&LayInNum&"</font>题，无法导入(这些题目不在任一类别中)：<font color=red>"&(Total-LayInNum)&"</font>题"
		call suc()
	end if
	tconn.close
	set tconn=nothing				
end sub
%>