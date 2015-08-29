<!--#include file="conn.asp"-->
<!--#include file="Z_TESTCONN.ASP"-->
<!-- #include file="inc/const.asp" -->
<% 
'-----------------
'edit by 我来了
'-----------------
	
	stats="开心辞典"

	if not founduser then
		call nav()
		call head_var(2,0,"","")	
		Errmsg=Errmsg+"<br>"+"<li>您没有进入开心辞典的权限，请先登录或者同管理员联系。"
		call dvbbs_error()
	else
		dim ars
		if request("action")="PaiXing" then
			stats="开心辞典 答题龙虎榜"
			call nav()
			call head_var(0,0,"开心辞典","Z_test.asp")			
			call PaiXing()
			call activeonline()
		else 
			call nav()
			call head_var(2,0,"","")
			call activeonline()		
			call main()
		end if		
	end if
	aconn.close
	set aconn=nothing
	call footer()
'===================================
sub main()
	dim aup,apass,acount,aokcount,auserjl,auserp,ausere,auserc,auserarticle
	dim auserans,auserok,auserl,adttime,azql,adts,afjjl,ats,ausersinew,auserupsinew
	dim AnswerSetting,OtherSetting,kxcd_ad

	session("checkpage")=""
	session("ajj")=""
	session("atitle")=""
	session("akey1")=""
	session("akey1")=""
	session("akey2")=""
	session("akey3")=""
	session("akey4")=""
	session("aok")=""
	session("aid")=""
	session("testlb")=""
	session("atime")=""
	session("atx")=""
	session("acount")=""
	session("aokcount")=""

	set ars=server.createobject("adodb.recordset")
	
	sql="select AnswerSetting,OtherSetting,userjl,zql,dts,fjjl,userarticle,userupsinew,ad from [config]"
	set ars=aconn.execute(sql)
	AnswerSetting=split(ars(0),"||")	'0每回答一题所需游戏币个数 ||1每回答一题魅力值增减值||2每回答一题经验值增减值||3两次玩之间的时间间隔||4答题数||5是否显示正确答案||6花钱买知识所需金钱||7答题是否限时||8题目是否支持UBB
	if ubound(AnswerSetting)<8 then
		redim preserve AnswerSetting(8)
		AnswerSetting(8)="1"
	end if
	OtherSetting=split(ars(1),"||")		'0是否显示图片 || 1图片路径 || 2答题龙虎榜显示记录条数
	auserjl=ars(2)						'上传奖励金额 
	azql=ars(3)							'获得附加奖励条件 当次答题正确率
	adts=ars(4)							'获得附加奖励条件 当次最少答题数
	afjjl=ars(5)						'附加奖励金额
	auserarticle=ars(6)					'每发一帖奖游戏币
	auserupsinew=ars(7)					'上传一题奖游戏币 
	kxcd_ad=ars(8)  					'广告代码
	ars.close  
	
	if request("action")="TestOut" and session("oklu")<>"" then
		call TestInfo(azql,adts,afjjl)
		exit sub
	end if	
	'response.Write "now()="&now()&"<br>DateAdd(""n"",-AnswerSetting(3),now())="&DateAdd("n",-AnswerSetting(3),now())
	sql="select userup,userpass,userans,userok,usersinew,ftl from [testuser] where username='"&membername&"'"
	set ars=aconn.execute(sql)
	if ars.eof or ars.bof then
		dim TempDate
		TempDate=DateAdd("n",-AnswerSetting(3),now())
		aconn.execute("insert into [testuser] (username,intime,ftl,usersinew,js) values('"&membername&"',#"&year(TempDate)&"-"&month(TempDate)&"-"&day(TempDate)&" "&hour(TempDate)&":"&minute(TempDate)&":"&second(TempDate)&"#,"&myarticle&","&auserarticle*myarticle&","&AnswerSetting(4)&")")
		aup=0:		apass=0:		auserans=0:		auserok=0:		auserl=0
	else
		aup=ars(0)
		apass=ars(1)
		auserans=ars(2)
		auserok=ars(3)
		ausersinew=ars(4)+(myarticle-ars(5))*auserarticle		'游戏币个数,帖子被删除，游戏币个数也会减少
		if myarticle<>ars(5) then aconn.execute("update [testuser] set ftl="&myarticle&",usersinew="&ausersinew&" where username='"&membername&"'")

		if auserans=0 then
			auserl=0
		else
			auserl=round(auserok/auserans,3)*100
		end if
	end if
	ars.close
%>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
  <tr><th colSpan="3" height=25>开心辞典----知识问答</th></tr>
  <tr>
    <td width="200" align="center" valign="top" class=tablebody2>
	<table align="center" cellPadding="3" cellSpacing="1" class=tableborder1 style="width:97%">
        <tr><th colSpan="2">各类题目详细列表</th></tr>
		  <tr>
				<td align="center" width="60%" Class=TopLighNav1><b>题目类型</b></td> 
				<td align="center" width="40%" Class=TopLighNav1>题量</td> 
		  </tr>
<%
			dim Cnum	'题库总数
			Cnum=0
			sql="select id,lb from testlb" 
			ars.open sql,aconn,1,1
             if ars.eof and ars.bof then
				response.Write "<tr><td colspan=2 height=32 class=tablebody1><font color=red>还没有任何分类和题目</font></td></tr>"
			else
				dim temp
				do while not ars.eof
					temp=aconn.execute("select count(id) from [test] where lb="&ars(0))(0)
%>
					<tr>
						<td align="center" width="60%" height="20" class=tablebody1><%if temp<>0 then%><a href="Z_Testii.asp?testlb=<%=ars(0)%>&btest=开始答题&action=dt" title="点击进入知识问答 "><%=ars(1)%></a><%else%><a title="该类别还没有任何题目" class="cgray"><%=ars(1)%></a><%end if%></td>
						<td align="center" width="40%" class=tablebody1><%=temp%></td>
					</tr>
<%				
					Cnum=Cnum+temp	
					ars.movenext
				loop
			end if
			ars.close
%>
		</table>
		<br><font color=<%=Forum_body(8)%>>点击题目类型可以直接进入知识问答</font>
		<br><hr width=180 size="1">
		题库总数：<font color=<%=Forum_body(8)%>><b><%=Cnum%></b></font>
		<%temp=aconn.execute("select count(id) from [testtemp]")(0)%>
		<br><%if temp>0 then%><a href="Z_TestAdminUpLoad.asp" title="立刻进入审核">未审核上传题</a><%else%>未审核上传题<%end if%>：<font color=<%=Forum_body(8)%>><b><%=temp%></b></font>		
		<hr width=180 size="1">
	</td>
    <td vAlign="middle" align="center" class=tableBody2>
	
	<table border=0><tr><td height="3"></td></tr></table>
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
		<tr>
			<td align=left valign=middle class=TopLighNav1>&nbsp;<a href="Z_test.asp?action=PaiXing">答题龙虎榜</a> | <span lang="zh-cn"><a href="Z_TestUpload.ASP">上载选择题</a></span> | <span lang="zh-cn"><a href="Z_TestUpload.ASP?action=addtk">上载填空题</a></span> | <a href="Z_testhelp.asp"><font color=red><b>问答帮助</b></font></td>
			<td align=center valign=middle class=TopLighNav1 width="60"><a href="z_testSetting.asp" onMouseOver='ShowMenu(Test_Adimn_Menu,60)'>辞典管理</a></td>  
		</tr> 
	</table>
	<table border=0><tr><td height="3"></td></tr></table>

	<table cellSpacing="1" cellPadding="3" width="100%" height="100%" class=tableBorder1 style="width:97%">
	<tr><td class=tablebody2>
		<table cellSpacing="1" cellPadding="0" width="100%" height="100%" class=tableBorder1 align="center" style="width:97%">
              <tr>
        <td align="center" colSpan="4" valign="middle" class=tablebody2 height=*>
<%
			if cint(OtherSetting(0))=1 then 
				response.Write "<img src="""&OtherSetting(1)&"""  width=""80"" height=""80"">"
			elseif cint(OtherSetting(0))=2 and kxcd_ad<>"" then
				response.Write kxcd_ad
			else	
				'set ars=aconn.execute("select top 10 Content,UserName,AddTime from testnews order by AddTime desc")
				'if ars.eof and ars.bof then
				'else
				'	response.Write "<marquee scrollamount=1 scrolldelay=5 width='98%' height=100 direction=up>" 
				'	do while not ars.eof
				'		response.Write ars(0)&"<br>"
				'		ars.movenext 
				'	loop				
				'	response.Write "</marquee>"					
				'end if
				'ars.close
			end if
%>
			</td></tr>
			<tr>
                <td valign="middle" class=tablebody1 height="140"> 
                  <table border=0 align=center cellspacing=1 class=tableBorder1 style="width:97%">
					  <tr bgcolor="#ffffff">
						<td align="left" colspan=2 class=TopLighNav1><b>◇您的属性◇</b></td>
					  </tr>
					  <tr bgcolor="#ffffff">
						<td align="right" width="50%">上传题数&nbsp;</td>
						<td width="50%"><font color="#FF0000">&nbsp;<b><%=aup%></b></font></td>
					  </tr>
					  <tr bgcolor="#ffffff">
						<td align="right">审核通过&nbsp;</td>
						<td><font color="#FF0000">&nbsp;<b><%=apass%></b></font></td>
					  </tr>		  
					  <tr bgcolor="#ffffff">
						<td align="right">答 题 数&nbsp;</td>
						<td><font color="#FF0000">&nbsp;<b><%=auserans%></b></font></td>
			
					  </tr>
					  <tr bgcolor="#ffffff">
						<td align="right">答 对 数&nbsp;</td>
						<td><font color="#FF0000">&nbsp;<b><%=auserok%></b></font></td>	
					  </tr>  
					  <tr bgcolor="#ffffff">
						<td align="right">正 确 率&nbsp;</td>
						<td><font color="#FF0000">&nbsp;<b><%=auserl%>%</b></font></td>
					  </tr>
					  <tr bgcolor="#ffffff">
						<td align="right">游 戏 币&nbsp;</td>
						<td><font color="#FF0000">&nbsp;<b><%=ausersinew%></b></font></td>
					  </tr>		  
				  </table>
				  
	  		    </td>
              </tr>
			</table>
	  				
		</td> 
		<td width=3 valign="top" class=tableBody1></td>
        <td valign="middle" align="center" width="200" class=tableBody2>
		
			<table width="100%" align=center cellspacing=1 class=tableBorder1 style="width:97%">
			  <tr bgcolor="#ffffff">
				<td align="LEFT" width="100%" class=TopLighNav1 colspan="2"><b>问答配置</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">时间间隔</td>
				<td><b><span lang="en-us">&nbsp;</span><font color="#FF0000"><%=AnswerSetting(3)%></font> </b>Mins/次</td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">取题目数</td> 
				<td><b><span lang="en-us">&nbsp;</span><font color="#FF0000"><%=AnswerSetting(4)%></font> </b>题/次</td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td class=TopLighNav1 colspan="2"><b>奖励配置</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">游戏币</td> 
				<td>&nbsp;每答一题<b><span lang="en-us">&nbsp;</span><font color="#FF0000">－<%=AnswerSetting(0)%></font></b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">经验值</td>
				<td><span lang="en-us">&nbsp;</span>对<span lang="en-us"><font color="#FF0000"><b> +</b></font></span><b><font color="#FF0000"><%=AnswerSetting(2)%></font></b> 错<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(2)%></font></b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">魅力值</td>
				<td><span lang="en-us">&nbsp;</span>对<b><font color="#FF0000"><span lang="en-us"> +</span><%=AnswerSetting(1)%></font></b> 错<b><font color="#FF0000"><span lang="en-us"> -</span><%=AnswerSetting(1)%></font></b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">现金</td>
				<td><span lang="en-us">&nbsp;</span>按题目不同增减</td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">附加奖励</td>
				<td>&nbsp;<b><font color=red><%=afjjl%></font>元</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">正确率</td>
				<td>&nbsp;<b><font color=red><%=azql%></font>%</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">最少答题</td>
				<td>&nbsp;<b><font color=red><%=adts%></font>题</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="left" class=TopLighNav1 colspan="2"><b>上载奖励</b></td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">成功收录</td>
				<td><b><font color="#cc3300"><span lang="en-us">&nbsp;</span><%=auserjl%><span lang="zh-cn">元</span></font></b>/题</td>
			  </tr>
			  <tr bgcolor="#ffffff">
				<td align="right">游戏币</td>
				<td>&nbsp;<b><font color=red><%=auserupsinew%></font>枚</b></td>
			  </tr>
			  <tr bgColor="#ffffff">
				<td colSpan="2">用户每在论坛发一帖奖游戏币&nbsp;<b><font color=red><%=auserarticle%></font>枚</b></td>
			  </tr>
			</table>
			
        </td>
      </tr>
    </table>
	
<!--	
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
		<tr><td align=left valign=middle class=TopLighNav1>管理选项： <a href="z_testSetting.asp">辞典设置</a> | <a href="Z_TestAdminUpLoad.asp">上传题目管理</a> | <a href="Z_TestEdit.asp">已有题库管理</a></td></tr> 
	</table>
	<br>
-->		
<%
'我来了修改 
 	sql="select username,UserHidden from online where boardid=0 and stats like '"&stats&"%' order by lastimebk" 
	ars.open sql,conn,1,1
%>
	<table border=0><tr><td height="3"></td></tr></table>
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
		<tr>
			<td align=left valign=middle class=tablebody2 width="115">开心辞典在线(<font color=red><%=ars.recordcount%>人</font>)</td>
			<td align=center valign=middle class=tablebody1><marquee scrollamount=1 scrolldelay=2 width='98%'> 
	<%			do while not ars.eof
					if cint(ars(1))=1 then
						if master or ars(0)=membername then
							response.Write ars(0)&"&nbsp;&nbsp;"
						else
							response.Write "隐身会员&nbsp;&nbsp;"	
						end if
					else
						response.Write ars(0)&"&nbsp;&nbsp;"
					end if		
					ars.movenext 
				loop				
	%>		
			</marquee></td>
		</tr> 
	</table>
<%
	ars.close
%>
    </td>
  </tr>
</table>
<%
end sub

sub PaiXing()
'----------------
' write by 我来了
'---------------- 
	dim PaiXingTop	'排行榜列出的记录个数
	dim orders,ordername
	sql="select OtherSetting from [config]"
	set ars=aconn.execute(sql)
	if ars.eof then	
		PaiXingTop=20
	else 
		PaiXingTop=split(ars(0),"||")(2)
		if not isInteger(PaiXingTop) then PaiXingTop=20
	end if	
	if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	else
		orders=cint(request("orders"))
	end if
	select case orders
	  case 1
		sql=" order by userok desc"
		ordername="答对数 排行TOP "&PaiXingTop			  
	  case 2
		sql=" order by userans desc"
		ordername="答题数 排行TOP "&PaiXingTop				  
	  case 3
		sql=" order by userup desc"
		ordername="上存题目数 排行TOP "&PaiXingTop
	  case 4
		sql=" order by userpass desc"
		ordername="上存通过率 排行TOP "&PaiXingTop				  
	  case 5
		sql=" order by usersinew desc"
		ordername="游戏币 排行TOP "&PaiXingTop
	  case 6
		sql=" order by intime desc"
		ordername="最后答题时间 排行TOP "&PaiXingTop		
	end select
%>
    　<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
      <tr><th height="25"><a name=mm>开心辞典排行榜</a></th></tr>
      <tr><td valign="top" class=tablebody1>
        <table border=0 width="100%" align=center cellspacing=1 height="78" class=tableBorder1>
        <tr>
            <td align="center" width="100%" height="25" colspan="7" class=tablebody1><font color=<%=Forum_body(8)%>><%=ordername%></font> (点击不同的项目可以按照不同的方式排序)</td>
          </tr>
        <tr>
            <td class=TopLighNav1 height="21" align="center">用户名</td> 
			<tD class=TopLighNav1 align="center"><%if orders=2 then%><font color=<%=Forum_body(8)%>>答题总数</font><%else%><a href=Z_test.asp?action=PaiXing&orders=2>答题总数</a><%end if%></td>  
			<td class=TopLighNav1 align="center"><%if orders=1 then%><font color=<%=Forum_body(8)%>>答对数</font><%else%><a href=Z_test.asp?action=PaiXing&orders=1>答对数</a><%end if%></td> 
			<td class=TopLighNav1 align="center"><%if orders=3 then%><font color=<%=Forum_body(8)%>>上存题目数</font><%else%><a href=Z_test.asp?action=PaiXing&orders=3>上存题目数</a><%end if%></td> 
			<td class=TopLighNav1 align="center"><%if orders=4 then%><font color=<%=Forum_body(8)%>>已采用的题目数</font><%else%><a href=Z_test.asp?action=PaiXing&orders=4>已采用的题目数</a><%end if%></td>  
			<td class=TopLighNav1 align="center"><%if orders=6 then%><font color=<%=Forum_body(8)%>>最后答题时间</font><%else%><a href=Z_test.asp?action=PaiXing&orders=6>最后时间</a><%end if%></td>
			<td class=TopLighNav1 align="center"><%if orders=5 then%><font color=<%=Forum_body(8)%>>游戏币</font><%else%><a href=Z_test.asp?action=PaiXing&orders=5>游戏币</a><%end if%></td>
        </tr>
<%			  set rs=server.createobject("adodb.recordset")	
			  sql="select top "&PaiXingTop&" username,userans,userok,userup,userpass,intime,usersinew from testuser "&sql
              rs.open sql,aconn,1,1
             if rs.eof and rs.bof then
%>
					<tr> 
					<td colspan=7 height="32" class=tablebody2><font color="red">　　还没有任何用户数据。</font></td>
					</tr>
<%
			else 
				dim II
              	II=0 
              	do while (not rs.eof)  
%>
				<tr>
					<td align="center" class=tablebody2><%=rs(0)%></td>
					<td align="center" class=tablebody1><%=rs(1)%></td>
					<td align="center" class=tablebody2><%=rs(2)%></td>
					<td align="center" class=tablebody1><%=rs(3)%></td>
					<td align="center" class=tablebody2><%=rs(4)%></td>
					<td align="center" class=tablebody1><%=rs(5)%></td>
					<td align="center" class=tablebody2><%=rs(6)%></td>
				  </tr>
<%
		  		rs.movenext
              	II=II+1
				if II>=clng(PaiXingTop) then exit do
              loop
           end if
		   rs.close
		   set rs=nothing
%>
		  </table></td>
      </tr>
	  <tr><td class=tablebody1 align="center">◇ &nbsp;<a href="Z_test.asp">返回辞典首页</a> &nbsp;◇</td></tr>
    </table>
<%
end sub

sub TestInfo(Rate,DTnum,JJ)  
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style=""width:75%"">"&_
				"<tr><th height=25>本次答题情况</th></tr>"&_
				"<tr><td width=""100%"" class=tablebody1 algin=center height=1>"

	 if session("oklu")>=Rate then
		if session("ok")>=DTnum then
			conn.execute ("update [user] set userWealth=userWealth+"&JJ&" where username='"&membername&"'")
			response.write "<br>&nbsp;&nbsp;你本次答题的正确率为：<font color=red>"&session("oklu")&"%</font>&nbsp;&nbsp;答对 <font color=red>"&session("ok")&"</font> 题<br><br>"
			response.write "&nbsp;&nbsp;达到附加奖励的标准，已将附加奖励<font color=red>"&JJ&"</font>两加入了你的现金帐户，请查收！<br><br>"
			session("oklu")=""
			session("ok")=""						
		else
			response.write "<br>&nbsp;&nbsp;你本次答题的正确率为：<font color=red>"&session("oklu")&"%</font>&nbsp;&nbsp;答对 <font color=red>"&session("ok")&"</font> 题<br><br>"
			response.write "&nbsp;&nbsp;由于你的答对数不足<font color=red>"&DTnum&"</font>题，所以附加奖金与你无缘了！<br><br>"
			session("oklu")=""
			session("ok")=""
		end if
	else
		response.write "<br>&nbsp;&nbsp;你本次答题的正确率为：<font color=red>"&session("oklu")&"%</font>&nbsp;&nbsp;答对 <font color=red>"&session("ok")&"</font> 题<br><br>"
		response.write "&nbsp;&nbsp;低于规定的最低正确率标准，所以附加奖金与你无缘了！<br><br>"
		session("oklu")=""
	end if
	response.write "</td></tr><tr><td align=center height=26 class=tablebody2><a href=z_test.asp>返回开心辞典</a></td></tr></table>"
end sub
%>
<script language="javascript">
var Test_Adimn_Menu= '<a style=font-size:9pt;line-height:14pt; href=\"z_testSetting.asp\">辞典设置管理</a><br><a style=font-size:9pt;line-height:14pt; href=\"Z_TestAdminUpLoad.asp\">上传题目管理</a><br><a style=font-size:9pt;line-height:14pt; href=\"Z_TestEdit.asp\">已有题库管理</a><br><a style=font-size:9pt;line-height:14pt; href=\"Z_TestAdminUser.asp\">辞典用户管理</a><br><a style=font-size:9pt;line-height:14pt; href=\"Z_TestLayIn.asp\">题库批量导入</a>'
</script>