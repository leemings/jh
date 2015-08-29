<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<!--#include file="z_gp_Chufa.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
stats="交易大厅"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if founderr then
	call dvbbs_error()
else
  call main()
end if
call activeonline()
call footer()

sub main()%>
<table cellspacing=0 cellpadding=0 align=center border=0 width="<%=Forum_body(12)%>">
  <tr> 
    <td width="25%"> 
    	<table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr> 
          <th height=25 align=center nowrap><%=Custom_Setting(0)%></th>
        </tr>
        <tr valign=middle> 
					<td height="54" class=tablebody2 align=center><%=Custom_Setting(1)%></td>
        </tr>
			</table>
		</td>
    <td width="50%"> 
    	<table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr> 
          <th height=25 align="center" nowrap>〓 <a href="z_gp_Announcements.asp">股市公告板</a> 〓</th>
        </tr>
        <tr> 
          <td height="54" class=tablebody2 align="center" id="content" style="FILTER: revealTrans(Transition=12, Duration=2);"> 
            <SCRIPT LANGUAGE='JavaScript' SRC='inc/z_gp_ad.js' TYPE='text/javascript'></script> 
						<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
							messages = new Array()
							mescolor = new Array()
							<%set rs=gp_conn.execute("select top 5 * from StockNews order by id desc")
							if rs.bof and rs.eof then
								%>messages[0]="欢迎光临股票交易所"
								mescolor[0]="blue"
								messages[1]="证监会提醒：股市风险莫测，谨慎入市！"
								mescolor[1]="red"
								<%else
								i=0
								do while not rs.eof 
									%>messages[<%=i%>]="<% response.Write rs("Title")&" <font color=gray>("&rs("AddTime")&")</font><br>"%>"
									mescolor[<%=i%>]="<% response.Write rs("color")%>"
									<%i=i+1
									rs.movenext
								loop
								%>messages[<%=i%>]="证监会提醒：股市风险莫测，谨慎入市！"
								mescolor[<%=i%>]="blue" 
							<%end if
							rs.close%>
							dotransition()
						</SCRIPT>
					</td>
        </tr>
      </table></td>
    <td width="25%"> 
	    <table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
	      <tr> 
	        <th colspan="2" height=25 align=center nowrap>电子挂历</th>
	      </tr>
	      <tr height="54" bgcolor="#000000" valign=middle> 
					<td>
						<table width="100%" border="0">
							<tr>
								<td align="left" valign="middle" nowrap rowspan=2><img src="images/img_gupiao/clock/cb.gif" name="a"><img src="images/img_gupiao/clock/cb.gif" name="b"><img src="images/img_gupiao/clock/colon.gif" name="c"><img src="images/img_gupiao/clock/cb.gif" name="d"><img src="images/img_gupiao/clock/cb.gif" name="e"><img src="images/img_gupiao/clock/colon.gif" name="f"><img src="images/img_gupiao/clock/cb.gif" name="g"><img src="images/img_gupiao/clock/cb.gif" name="h"><img src="images/img_gupiao/clock/cam.gif" name="j"></td>
								<td align="center" valign="middle" nowrap><font color="#FFFFFF"><%=WeekdayName(Weekday(Date))%></font></td>
							</tr>
							<tr>
	            	<td align="center" nowrap><font color="#FFFFFF"><%=DatePart("yyyy", Date) & " 年 " & DatePart("m", Date) & " 月 " & DatePart("d", Date) & " 日"%></font>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
		</tr>
</table><br>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
  <tr> 
    <th align="center" height=25 nowrap>股票名称</th>
    <th align="center" height=25 nowrap>经营者</th>
    <th align="center" height=25 nowrap>交易量</th>
    <th align="center" height=25 nowrap>昨收盘价</th>
    <th align="center" height=25 nowrap>成交价</th>
    <th align="center" height=25 nowrap>买入笔数</th>
    <th align="center" height=25 nowrap>卖出笔数</th>
    <th align="center" height=25 nowrap>成交量</th>
    <th align="center" height=25 nowrap>涨跌幅</th>
  </tr>
<%  
	dim CurrentPrice,ZangDieRange,sstop	'股票当前价格，股票涨跌幅度，股票状态(停牌，涨停，跌停等)
	Dim ColExp 		'综合指数
	Dim color
	stats="股市大厅"
	sql= "select * from [GuPiao] order by sid"         
	set rs=gp_conn.execute(sql)         
	DO while not RS.EOF  
		if rs("DangQianJiaGe")<=1 then           
			color="#808080"           
		elseif rs("KaiPanJiaGe")<rs("DangQianJiaGe") then           
			color="#FF0000"           
		elseif rs("KaiPanJiaGe")>rs("DangQianJiaGe") then          
			color="#0000FF"           
		elseif rs("KaiPanJiaGe")=rs("DangQianJiaGe") then             
			color="#000000	"          
		end if          
		ZangDieRange=round((rs("DangQianJiaGe")-rs("KaiPanJiaGe"))/rs("KaiPanJiaGe"),4)
		CurrentPrice=round(rs("DangQianJiaGe"),2)
		if CurrentPrice<=Trade_Setting(11)+0 then
			sstop=4
			ZangDieRange="<b>停牌</b>"
		elseif ZangDieRange>=Trade_Setting(9)+0 then          
			sstop=1
			ZangDieRange="<b>涨停</b>"          
		elseif ZangDieRange<=-Trade_Setting(10)+0 then 
			sstop=2 
			ZangDieRange="<b>跌停</b>" 
		elseif ZangDieRange=0 then
			sstop=3
			ZangDieRange="0.00 %"
		else       
			sstop=0
		end if   
%>
  <tr valign="middle"> 
    <td class=tablebody1 height="20" width="14%" align="center" nowrap><a href="z_gp_DispCompare.asp?sid=<%=rs("sid")%>"><font color="<%=color%>"><%=rs("QiYe")%></font></a></td>
    <td class=tablebody2 height="20" align="center" nowrap><%=iif(rs("JingYingZhe")="","<font color="&color&">-</font>","<a href='z_gp_Dispu.asp?Uid="&rs("uid")&"'><font color="&color&">"&rs("JingYingZhe")&"</font></a>")%></td>
		<td class=tablebody1 height="20" align="right" nowrap><font color="<%=color%>"><%=rs("JiaoYiLiang")%></font>&nbsp;</td>
    <td class=tablebody2 height="20" align="right" nowrap><font color="<%=color%>"><%=formatnumber(rs("KaiPanJiaGe"),2,true)%></font>&nbsp;</td>
    <td class=tablebody1 height="20" align="right" nowrap><font color="<%=color%>"><%=formatnumber(rs("DangQianJiaGe"),2,true)%></font>&nbsp;</td>
    <td class=tablebody2 height="20" align="right" nowrap><font color="<%=color%>"><%=iif(rs("MaiRuBiShu")=0,"0",rs("MaiRuBiShu"))%></font>&nbsp;</td>
		<td class=tablebody1 height="20" align="right" nowrap><font color="<%=color%>"><%=iif(rs("MaiChuBiShu")=0,"0",rs("MaiChuBiShu"))%></font>&nbsp;</td>
		<td class=tablebody2 height="20" align="right" nowrap><font color="<%=color%>"><%=iif(rs("ChengJiaoLiang")=0,"0",rs("ChengJiaoLiang"))%></font>&nbsp;</td>
    <td class=tablebody1 height="20" align="right" nowrap><font color="<%=color%>"><% 
			if sstop<>0 then
				response.write ""&ZangDieRange&""
			else
				response.write ""&formatnumber(ZangDieRange*100,2,true)&" %"
				ColExp=ColExp+ZangDieRange*100		'计算综合指数
			end if 
		%></font>&nbsp;</td>
  </tr>
  <%        
		rs.MoveNext         
	Loop 
	rs.close
	set rs=Nothing
%>
</table>
<%
'==============个人股票管理=============
	if HaveAccount then
%>
	<br>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
  <TR> 
    <Td colSpan=10 height=25 class=tablebody2 align=center>〓 <font color=<%=forum_body(8)%>>个股管理</font> 〓&nbsp;&nbsp;欢迎您，<b><%=membername%></b>，您的持股情况如下：</Td>
  </TR>
  <TR align=center valign="middle" > 
    <th nowrap>股票名称</Th>
    <th nowrap>拥有数量</Th>
    <th nowrap>持股成本</Th>
    <th nowrap>当前价</Th>
    <th nowrap>持股价</Th>
    <th nowrap>持股现值</Th>
    <th nowrap>每股盈亏</Th>
    <th nowrap>总盈亏</Th>
    <th nowrap>盈亏率</Th>
	  </TR>
<%
	Dim rst,AgvInPrice		'某股票当前价格 ,某股票平均买入价格   
	Dim MyGupiaoNum,MyGupiaoValue,MyGupiaoCost 		'个人持股总数，个人股票市值,个人股票成本             
	Dim YKStr,TotalGranRate  '盈亏情况,总盈亏率
	MyGupiaoNum=0 :	MyGupiaoValue=0 :	MyGupiaoCost=0
	sql="select d.*,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&MyUserID&" order by d.[ChiGuShu] desc"  
	set rs=gp_conn.execute(sql)
	if rs.eof and rs.bof then
		RESPONSE.Write "<tr><td colspan=10 class=tablebody1 align=center><b>你现在没有股票</b></tr>"
	else
		do while not rs.EOF   
			CurrentPrice=round(rs("DangQianJiaGe"),2)
			AgvInPrice=round(rs("PingJunJiaGe"),2)
			MyGupiaoNum=MyGupiaoNum+rs("ChiGuShu")
			MyGupiaoValue=MyGupiaoValue+rs("ChiGuShu")*CurrentPrice
			MyGupiaoCost=MyGupiaoCost+rs("ChiGuShu")*AgvInPrice	
%>
		<TR align="right"> 
			<TD class=tablebody1 nowrap align=center><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></TD>
			<TD class=tablebody1 nowrap><%=rs("ChiGuShu")%>&nbsp;</TD>
			<TD class=tablebody1 nowrap><%=formatnumber(AgvInPrice*rs("ChiGuShu"),2,True)%>&nbsp;</TD>
			<TD class=tablebody1 nowrap><%=formatnumber(CurrentPrice,2,True)%>&nbsp;</TD>
			<TD class=tablebody1 nowrap><%=formatnumber(AgvInPrice,2,True)%>&nbsp;</TD>
			<TD class=tablebody1 nowrap><%=formatnumber(CurrentPrice*rs("ChiGuShu"),2,True)%>&nbsp;</TD>
			<TD class=tablebody1 nowrap><%
			dim bj,TotalGain,GranRate		'单股盈亏,单个股票总盈亏,单个股票总盈亏率
			bj=CurrentPrice-AgvInPrice
			bj=round(bj,2)
			if bj>0 then
				response.write "<font color=#FF0000>↑ "&formatnumber(bj,2,true)	&"</font>"
			elseif bj<0 then
				response.write "<font color=#0000FF>↓ "&formatnumber(abs(bj),2,True)&"</font>"
			elseif bj=0 then
				response.write "<font color=#000000>→ 0.00</font>"
			end if
			response.Write "&nbsp;</TD><TD class=tablebody1 nowrap>"
			
			TotalGain=bj*rs("ChiGuShu")
			GranRate=TotalGain*100/MyGupiaoCost
			if TotalGain>0 then 
				response.Write "<font color=#FF0000>↑ "&formatnumber(TotalGain,2,True)&"</font>&nbsp;</TD>" 
				response.Write "<TD class=tablebody2><font color=#FF0000>↑ "&formatnumber(GranRate,2,True)&" %</font>&nbsp;</TD></tr>"
			elseif TotalGain<0 then
				response.Write "<font color=#0000FF>↓ "&formatnumber(-TotalGain,2,True)&"</font>&nbsp;</TD>"
				response.Write "<TD class=tablebody1 nowrap><font color=#0000FF>↓ "&formatnumber(-GranRate,2,True)&" %</font>&nbsp;</TD></tr>" 
			else
				response.Write "<font color=#000000>→ 0.00</font>&nbsp;</TD>" 
				response.Write "<TD class=tablebody2><font color=#000000>→ 0.00 %</font>&nbsp;</TD></tr>"	
			end if	
			rs.MoveNext          
		Loop
		rs.close
	end if
	if MyGupiaoValue-MyGupiaoCost>0 then
		YKStr="<font color=red>↑ "&formatnumber(MyGupiaoValue-MyGupiaoCost,2,true)
	elseif MyGupiaoValue-MyGupiaoCost<0 then
		YKStr="<font color=#0000FF>↓ "&formatnumber(MyGupiaoCost-MyGupiaoValue,2,true)
	else
		YKStr="<font color=#000000>→ 0.00"
	end if
	YKStr=YKStr&"</font>" 
	if MyGupiaoCost=0 then
		TotalGranRate=0
	else	
		TotalGranRate=(MyGupiaoValue-MyGupiaoCost)*100/MyGupiaoCost
	end if	
%>	
  <TR align=right valign="middle"> 
    <Td height=20 class=tablebody2 align=center>合&nbsp;&nbsp;&nbsp;&nbsp;计</Td>
    <Td class=tablebody2><%=formatnumber(MyGupiaoNum,0,true)%>&nbsp;</Td>
    <Td class=tablebody2><%=formatnumber(MyGupiaoCost,2,true)%>&nbsp;</Td>
    <td class=tablebody2 align=center>-</Td>
    <Td class=tablebody2 align=center>-</Td>
    <Td class=tablebody2><%=formatnumber(MyGupiaoValue,2,true)%>&nbsp;</Td>
    <Td class=tablebody2 align=center>-</Td>
    <Td class=tablebody2><%=YKStr%>&nbsp;</Td>
    <Td class=tablebody2><%if TotalGranRate>0 then%><font color=#FF0000>↑ <%=formatnumber(TotalGranRate,2,true)%> %</font><%elseif TotalGranRate<0 then%><font color=#0000FF>↓ <%=formatnumber(-TotalGranRate,2,true)%> %</font><%else%>→ 0.00 %<%end if%>&nbsp;</Td>
		</TR>
	</TABLE>		
<%		
end if	
%> 
<br>
<table cellspacing=0 cellpadding=0 align=center border=0 width="<%=Forum_body(12)%>">
	<tr>
    <td valign=top width="30%"> 
      <!--第一列-->
      <table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr> 
          <th colspan="2" height=25 align=center>今日股市行情</th>
        </tr>
        <tr> 
          <td class=tablebody2>&nbsp;成交总额：</td>
          <td class=tablebody1 align=right><%=ChengJiaoMoney%> 元&nbsp;</td>
        </tr>
        <tr> 
          <td class=tablebody2>&nbsp;成交笔数：</td>
          <td class=tablebody1 align=right><%=ChengJiaoNum%> 笔&nbsp;</td>
        </tr>
        <tr> 
          <td class=tablebody2>&nbsp;综合指数：</td>
          <td class=tablebody1 align=right><%
          	if ColExp>0 then
          		%><font color=#FF0000>↑<%
          	elseif ColExp<0 then
          		%><font color="#0000FF">↓<%
          	else
          		%><font color="#000000">→<%
          	end if
          %> <%=formatnumber(abs(ColExp),2,true)%> %</font>&nbsp;</td>
				</tr>
      </table>
      <br> 
      <table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr> 
          <th colspan="2" height=25 align=center> 个人情况总揽</th>
        </tr>
        <%if HaveAccount then%>
          <tr> 
            <td class=tablebody2>&nbsp;帐户资金：</td>
            <td class=tablebody1 align=right><%=formatnumber(MyCash,2,true)%> 元&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;持股总数：</td>
            <td class=tablebody1 align=right><%=formatnumber(MyGupiaoNum,0,true)%> 股&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;持股市值：</td>
            <td class=tablebody1 align=right><%=formatnumber(MyGupiaoValue,2,true)%> 元&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;股票成本：</td>
            <td class=tablebody1 align=right><%=formatnumber(MyGupiaoCost,2,true)%> 元&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;盈亏情况：</td>
            <td class=tablebody1 align=right><%=YKStr%> 元&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;资产合计：</td>
            <td class=tablebody1 align=right><%=formatnumber(MyGupiaoValue+MyCash,2,true)%> 元&nbsp;</td>
          </tr>
          <tr> 
            <td class=tablebody2>&nbsp;今日买卖：</td>
            <td class=tablebody1 align=right><%=MyBiShu%> 笔&nbsp;</td>
          </tr>
	        <tr> 
	          <th colspan="2" align=center height=21><a href="z_gp_Mymanage.asp" class="cblue">账户管理</a></th>
	        </tr>
				<% else %>
					<form name="form1" method="post" action="z_gp_MyManage.asp?action=new">
          <tr>
            <td class=tablebody2 colspan="2"> <input type=hidden name=username value="<%=session("uname")%>"><br> &nbsp;&nbsp;&nbsp;&nbsp;你在股票市场里还没有帐户，要买卖股票请先开户。<br><br></td>
          </tr>
          <tr> 
            <th align=center colspan="2"><input type="submit" name="submit" value="开户" ></th>
          </tr>
					</form>
				<% end if %>	
      </table>
		</td>
    <!--第一列和第二列之间的分隔列-->
    <td width="40%" align=right valign=top> 
	    <!-- 第二列--股市实时新闻  -->
			<table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr> 
          <th height=25 align=center>股市实时新闻</td>
        </tr>
        <tr bgcolor="#000000"> 
          <td align="center"><MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 scrollDelay=10 direction=up width="100%" height=232 align="left" border="0"><%
						set rs=server.createobject("adodb.recordset")
						sql="select content,Addtime from RndEvent order by ID desc"
						rs.open sql,gp_conn
						do while Not RS.Eof
							response.Write ""&rs(0)&"<br><font color=""#FFCC99""><div align=""right"">"&rs(1)&"</div></font>"
							RS.MoveNext
						Loop
						rs.close
					%></marquee></td>
        </tr>
        <tr>
          <th align=center height=25><a href="z_gp_News.asp" class="cblue" title="查看所有新闻">更多新闻</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="z_gp_PStockApply.asp" class="cblue">个股上市</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="index.asp" class="cblue">返回论坛</a></th>
        </tr>
      </table>
		</td>
    <!--第二列和第三列之间的分隔列-->
    <td width="30%" align="right" valign="top"> 
      <!--第三列-->
      <table cellspacing=1 cellpadding=2 align="center" bgcolor=<%=Forum_Body(27)%> width=97%>
        <tr align=center> 
          <th colspan="3" height=25>股市总资产排行榜</td>
        </tr>
				<%set rs=gp_conn.execute("select top 12 ZhangHao,ZongZiJin,id from [KeHu] where suoding<2 order by [ZongZiJin] desc")
				if RS.Eof then 
					response.Write "<td class=tablebody2 colspan=4> 还未有人买入股票 </td>"
				else 
					i=0
					do while Not RS.Eof
						i=i+1
						response.Write "<tr><td class=tablebody1 align=""center"" width=""40""><font color=red>"&i&"</font></td>"	
						response.Write "<td class=tablebody2 align=""center""><a href='z_gp_Dispu.asp?uid="&rs(2)&"&username="&rs(0)&"'>"&rs(0)&"</a></td>"
						response.write "<td class=tablebody1 align=right>"
						if rs(1)<=0 then 
							response.Write "0 元"
						else
							response.Write formatnumber(rs(1),0,true)&" 元"
						end if
						response.Write "&nbsp;</td></tr>"
						RS.MoveNext
						if i>=12 then exit do
					Loop
					rs.close 
				end if%>
        <tr> 
          <th colspan=3 align=center height=21 valign="middle"><a href="z_gp_TopUser.asp?action=Reflash" class="cblue">刷新排名</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="z_gp_TopUser.asp" class="cblue">详细排名</a></th>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<% if master then %>
	  <tr> 
	    <th height=25>管理选项</th>
	  </tr>
	  <tr> 
	    <td class=tablebody2 height=25 valign="middle" align=center><font color="#0066FF">〓 <a href="z_gp_Admin_Price.asp?action=UpPrice" title="提升所有股票的现价">全红</a> ｜ <a href="z_gp_Admin_Price.asp?action=DownPrice" title="降低所有股票的现价">全绿</a> ｜ <a href="z_gp_Admin_Setting.asp">基本设置</a> ｜ <a href="z_gp_Admin_Gupiao.asp">股票管理</a> ｜ <a href="z_gp_Admin_Gupiao.asp?action=newgupiao">新股上市</a> ｜ <a href="z_gp_Admin_User.asp">客户管理</a> ｜ <a href="z_gp_Admin_Data.asp">数据库管理</a> 〓</font></td>
	  </tr>
	<%end if%>  
  <tr valign="bottom"> 
    <td class=tablebody1 id="clock" align=center>欢迎你</td>
  </tr>
</table>
<!--#include file="inc/z_gp_js.asp"-->
<%end sub%>
