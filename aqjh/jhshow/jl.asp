<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
on error resume next
aqjh_name=Session("aqjh_name")
search=cstr(server.HTMLEncode(Request("search")))
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
%>
<html><title>江湖秀大赛列表</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0" background=../bg.gif>
<div align="center"><font size="+3" face="黑体"><strong>江湖秀大赛</strong></font> 
  <br>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
    <tr>
      <td width="26%" height="201" align="center" valign="top">
<table width="98%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
          <tr> 
            <td align="center"> <font color="#0000FF"><strong> 
              <%ndate=Day(date())
if ndate<=20 then%>
              江湖秀大赛报名 
              <%elseif ndate<=27 then%>
              江湖秀大赛评选排行 
              <%else%>
              江湖秀大赛奖品领取 
              <%end if%>
              </strong></font> </td>
          </tr>
          <tr> 
            <td align="center"><font color="#0000FF"><strong> 
<%ndate=Day(date())
'在27号招,当第一个打开时判断
if ndate>27 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from pool", conn, 1,1
	if (isnull(rs("p_男第一名")) or rs("p_男第一名")="") and (isnull(rs("p_男第二名")) or rs("p_男第二名")="") and (isnull(rs("p_女第一名")) or rs("p_女第一名")="") and (isnull(rs("p_女第二名")) or rs("p_女第二名")="") then
		rs.close
		rs.Open "select * from sai where s_性别='男' Order by s_票数 DESC,s_加入时间", conn, 1,1
		if RS.Eof then
			nan1="无"
			nanp1=0
			nan2="无"
			nanp2=0
		else
			nan1=rs("s_姓名")
			nanp1=rs("s_票数")
			RS.MoveNext
			nan2=rs("s_姓名")
			nanp2=rs("s_票数")
		end if
		rs.close
		rs.Open "select * from sai where s_性别='女' Order by s_票数 DESC,s_加入时间", conn, 1,1
		if RS.Eof then
			nv1="无"
			nvp1=0
			nv2="无"
			nvp2=0
		else
			nv1=rs("s_姓名")
			nvp1=rs("s_票数")
			RS.MoveNext
			nv2=rs("s_姓名")
			nvp2=rs("s_票数")
		end if
		rs.close
		rs.Open "select * from pool", conn, 1,3
		rs("p_男第一名")=nan1
		rs("p_男第一票数")=nanp1
		rs("p_男第一领奖")=false
		rs("p_男第二名")=nan2
		rs("p_男第二票数")=nanp2
		rs("p_男第二领奖")=false
		rs("p_女第一名")=nv1
		rs("p_女第一票数")=nvp1
		rs("p_女第一领奖")=false
		rs("p_女第二名")=nv2
		rs("p_女第二票数")=nvp2
		rs("p_女第二领奖")=false
		rs.update
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
if ndate<=20 then%>
</strong><font color="#FF0000">参加大赛拿大奖，我要<a href="#" onClick="window.open('biaoming.asp','biaoming','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=320,height=375')" ><strong>报名</strong></a>，投票时间从<strong>21</strong>号开始!</font><strong> 
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,3
if rs("p_男第一名")<>"" or rs("p_男第二名")<>"" or rs("p_女第一名")<>"" or rs("p_女第二名")<>"" then
	rs("p_投票人")=""
	rs("p_投票人数")=0
	rs("p_金币")=0
	rs("p_报名人数")=0
	rs("p_男第一名")=""
	rs("p_男第一票数")=0
	rs("p_男第一领奖")=false
	rs("p_男第二名")=""
	rs("p_男第二票数")=0
	rs("p_男第二领奖")=false
	rs("p_女第一名")=""
	rs("p_女第一票数")=0
	rs("p_女第一领奖")=false
	rs("p_女第二名")=""
	rs("p_女第二票数")=0
	rs("p_女第二领奖")=false
	rs.update
	conn.execute("delete from sai")
end if

rs.close
rs.Open "select * from pool", conn, 1,3
%>
              <br>
              报名人数：<font color=red><%=rs("p_报名人数")%></font>名<br>
              金币累计：<font color=red><%=rs("p_金币")%></font>个<br>
              预计奖金：<font color=red><%=(int(rs("p_金币")*0.6))%></font>个 
              <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
              </strong></font> 
              <%elseif ndate<=27 then
%>
              <table width="100%" border="0" cellpadding="1" cellspacing="2" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
                <tr align="center" bgcolor="#FFCC00"> 
                  <td>报名ID</td>
                  <td>得票数</td>
                </tr>
                <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
xx=0
rs.Open "select * from sai where s_性别='男' Order by s_票数 DESC,s_加入时间", conn, 1,1
do while Not RS.Eof
%>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left" bgcolor="#FFCC00"><%=rs("s_姓名")%>(男)</td>
                  <td align="right" bgcolor="#FFCC00"><%=rs("s_票数")%></td>
                </tr>
                <%xx=xx+1
if xx>=5 then exit do
RS.MoveNext
loop
rs.close
xx=0
rs.Open "select * from sai where s_性别='女' Order by s_票数 DESC,s_加入时间", conn, 1,1
do while Not RS.Eof
%>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("s_姓名")%>(女)</td>
                  <td align="right" bgcolor="#FFCC00"><%=rs("s_票数")%></td>
                </tr>
                <%xx=xx+1
if xx>=5 then exit do
RS.MoveNext
loop%>
              </table>
              <%rs.close
              rs.Open "select * from pool", conn, 1,3
              baoren=replace(rs("p_投票人"),"|","<br>")
              %>
              报名人数：<font color=red><%=rs("p_报名人数")%></font>名<br>
              金币累计：<font color=red><%=rs("p_金币")%></font>个<br>
              预计奖金：<font color=red><%=(int(rs("p_金币")*0.6))%></font>个<br>
              投票人数：<font color=red><%=rs("p_投票人数")%></font>人 <br> <br> 
              <%rs.close
set rs=nothing
conn.close
set conn=nothing
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,1
 baoren=replace(rs("p_投票人"),"|","<br>")
%></strong></font>
              <table width="100%" border="0" cellpadding="2" cellspacing="2" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
                <tr align="center" bgcolor="#FFCC00"> 
                  <td width="49%">姓名</td>
                  <td colspan="2">操作</td>
                </tr>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_男第一名")%>(男1)</td>
                  <td width="27%" align="right"><%=rs("p_男第一票数")%></td>
                  <td width="24%"><%if rs("p_男第一领奖")=false then%><a href="lingjiang.asp?act=1">领奖</a><%else%>已领<%end if%></td>
                </tr>
                 <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_男第二名")%>(男2)</td>
                  <td width="27%" align="right"><%=rs("p_男第二票数")%></td>
                  <td width="24%"><%if rs("p_男第二领奖")=false then%><a href="lingjiang.asp?act=2">领奖</a><%else%>已领<%end if%></td>
                </tr>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_女第一名")%>(女1)</td>
                  <td width="27%" align="right"><%=rs("p_女第一票数")%></td>
                  <td width="24%"><%if rs("p_女第一领奖")=false then%><a href="lingjiang.asp?act=3">领奖</a><%else%>已领<%end if%></td>
                </tr>
                 <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_女第二名")%>(女2)</td>
                  <td width="27%" align="right"><%=rs("p_女第二票数")%></td>
                  <td width="24%"><%if rs("p_女第二领奖")=false then%><a href="lingjiang.asp?act=4">领奖</a><%else%>已领<%end if%></td>
                </tr>
              </table>
              <font color="#0000FF"><strong> 请排上名字的玩进入领取奬金！<br>
              报名人数：<font color=red><%=rs("p_报名人数")%></font>名<br>
              金币累计：<font color=red><%=rs("p_金币")%></font>个<br>
              预计奖金：<font color=red><%=(int(rs("p_金币")*0.6))%></font>个<br>
              投票人数：<font color=red><%=rs("p_投票人数")%></font>人 
              <%rs.close
              set rs=nothing
              conn.close
              set onn=nothing
              end if%>
              </strong></font></td>
          </tr>
          <tr> 
            <td align="center"><font color="#0000FF">比赛规则</font></td>
          </tr>
          <tr> 
            <td><strong>1.参赛方法？</strong><br>
              每次参赛将收取报名费<font color="#FF0000"><strong>10个金币</strong></font>，每人只限报名一次，所以你要慎重考虑参加报名，一但报名你的报名形将将不能再更改。 
              <br> <br> <strong>2.评选办法？</strong><br>
              江湖等级在<%=Application("aqjh_newplay")%>级以上的玩家可以参加评选择，参加投票的玩家每次系统赠送<strong><font color="#FF0000">1万银两</font></strong>作为鼓励！<br>
              每人每次可以投票一次，对于拉票的玩家我们一经查出删除ID处分。<br> <font color="#0000FF">男玩家投女玩这有的票</font>，<font color="#FF0000">女玩家投男玩家的票</font>!<br> 
              <br> <strong>3.奖励办法？</strong><br>
              每次我们将评选男、女各2名进行奖励，奖励的金额：<br> <font color="#0000FF">第一名为总报名费用的%20</font><br> 
              <font color="#0000FF">第二名为总报名费用的%10</font><br>
              20x2+10x2=60% 系统奖为总报名费用的60%!看看谁将是本次的幸运儿！<br> <strong><br>
              4.时间安排！</strong> <br> <font color="#0000FF">每月1-20号为报名时间</font>，<font color="#FF0000">21-27号为评选时间</font>,<font color="#0000FF">28号到月底为领奖品时间</font>，<strong>在领奖品时间没有领奖品的用户将会按弃权处理!<br>
              <br>
              5.注意事项！<br>
              </strong>a.相同名次按报名先后定名次。<br>
              b.因为是投票操作,有时难免有相同服装而不同票数，这一点请大家理解。<br>
              c.对于作弊，拉票行为一经发现马上取消参赛资格，对于严重者删除ID处份。</td>
          </tr>
          <%if ndate>=21 then%>
          <tr>
            <td><marquee scrollamount="1" scrolldelay="1" direction= "up"  height="350">
            <div align=center><%=baoren%></div>
            </marquee>
            </td>
          </tr>
          <%end if%>
        </table></td>
      <td width="74%" valign="top"> 
        <%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from sai  Order by s_加入时间", conn, 1,1
if rs.eof and rs.bof then%>
        <div align="center">对不起，暂时没有发现任何记录！！ <br>[ <a href=# onclick=history.go(-1)>返回上一级</a> ]<br></div>
<%else
RS.PageSize=20'设置每页记录数,可修改
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数
Dim Totalcount
Totalcount =INT(RS.recordcount)
StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount%>
<table border="1" width="100%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr>
            <td colspan=6 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff">江湖秀大赛选手名单(<font color="#0000FF"> 
              <%ndate=Day(date())
if ndate<=20 then%>
              按添加日期排序 
              <%elseif ndate<=27 then%>
              按添加日期排序 
              <%else%>
              按得票数排 
              <%end if%>
              排行不分男女</font>)</td>
          </tr>
<%I=0
p=RS.PageSize*(Curpage-1)
Response.Write"<tr>"
do while (Not RS.Eof) and (I<RS.PageSize)
	p=p+1
	myrnd=int(rnd*999)+1000
	Response.Write"<td   width=""25%"" height=""226"" align=center valign=""top"">"
	Response.Write"<div id=show"& myrnd &" style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0; HEIGHT: 226px;'></div>"
	Response.Write "<script Language=Javascript>var jhshow"& rs("s_id") &"='"& rs("s_数据") &"';"
	Response.Write "var showArray = jhshow"& rs("s_id") &".split('|');"
	Response.Write "var s='';"
	Response.Write "for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
	Response.Write "s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';"
	Response.Write "}}"
	Response.Write "s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &" title=""ID:"& rs("s_id") &" 得票："& rs("s_票数") &"&#13&#10说明："& rs("s_说明") &""">';"
	Response.Write "show"& myrnd &".innerHTML=s;</script>"
	if ndate<=27 and ndate>20 then
		Response.Write "<a href=""toupiao.asp?id="& rs("s_id") &""" ><font color=red><b>投票</b></font></a> "
		Response.Write "得票:"& rs("s_票数") &"<br>"
	end if
	Response.Write "<font color=blue>"& rs("s_说明") &"</font>"
	Response.Write"</td>"
	I=I+1
	if i/4=int(i/4) then Response.Write"</tr><tr>"
	RS.MoveNext
Loop
if i/4<>int(i/4) then Response.Write"</tr>"
%>
<tr>
            <td colspan=6 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
              有<font color=blue><%=Totalcount%></font>条 共<%=int(Totalcount/25)+1%>页 
              [ 
              <% if StartPageNum>1 then %>
              <a href="jl.asp?CurPage=<%=StartPageNum-1%>&fs=<%=fs%>"><<</a> 
              <%end if%>
              <% For I=StartPageNum to EndPageNum
if I<>CurPage then %>
              <a href="jl.asp?CurPage=<%=I%>&fs=<%=fs%>"><%=I%></a> 
              <%else %>
              <b><font color=#ff0000><%=I%></font></b> 
              <%end if %>
              <%Next %>
              <%if EndPageNum<RS.Pagecount then %>
              <a href="jl.asp?CurPage=<%=EndPageNum+1%>&fs=<%=fs%>">>></a> 
              <%end if%>
              ] 
              <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
            </td>
          </tr></table>
      </td>
    </tr></table>
</div></body></html>