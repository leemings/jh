<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<!--#include file="z_visual_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
stats="結婚教堂"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区结婚礼堂的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()%>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=762 align="center" height="1">
  <TBODY> 
  <TR>
    <TD bgColor=#FFFFFF height=1 vAlign=top width=180 rowspan="2">
    <img border="0" src="images/img_marry/01.jpg"><img border="0" src="images/img_marry/02.jpg"></TD>
    <TD bgColor=#FFFFFF vAlign=top width=39 rowspan="2" height="1">
    <img border="0" src="images/img_marry/03.jpg"></TD>
    <TD bgColor=#FFFFFF vAlign=top width=295 background="images/img_marry/04.jpg" height="1">
      　<p>　</p>
      <TABLE border=1 borderColor=#FFFFFF cellPadding=0 
      cellSpacing=0 width=100% style="border-collapse: collapse" align="left" height="1">
        <TR>
          <TD bordercolor="#FFFFFF" valign="top" align="justify" style="border-style: solid; border-width: 1" height="1">
            <div align="center">
              <center>
            <TABLE border=1 borderColor=#FFFFFF cellSpacing=1 width="100%" style="border-collapse: collapse" height="1">
              <TBODY> 
              <!--<TR>
                <TD width="15%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>求婚人</B></DIV></TD>
                <TD width="15%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>意中人</B></DIV></TD>
                <TD width="20%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>求婚时间</B></DIV></TD>	
                <TD width="20%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>爱的表白</B></DIV></TD>
                <TD width="12%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>礼物</B></DIV></TD>			  
                <TD width="18%" bordercolor="#FFFFFF" height="25"> 
                  <DIV align=center><B>操作</B></DIV></TD></TR>
				  	<% 
	set rs1=server.createobject("adodb.recordset")
    rs1.open"select * from qiuhun where dlg=false",connj,3,1
  	if rs1.bof then 
	rs1.close  %>
              <TD width=21% height="36" colspan=6>还没有任何人求婚呢</TD>
              </tr>  
              <%else 
                do while not rs1.eof %><TR>
                <TD width="15%" bordercolor="#FFFFFF" height="1">
                  <div align="center"><a href=dispuser.asp?name=<%=rs1("username")%> target="_blank" title="查看<%=rs1("username")%>的资料"><%=rs1("username")%></a></div>
                </TD>
                <TD width="15%" bordercolor="#FFFFFF" height="1">
                  <div align="center"><a href=dispuser.asp?name=<%=rs1("tousername")%> target="_blank" title="查看<%=rs1("tousername")%>的资料"><%=rs1("tousername")%></a></div>
                </TD>
                <TD width="20%" bordercolor="#FFFFFF" height="1">
                  <div align="center"><%=rs1("addtime")%></div>
                </TD>
                <TD width="20%" bordercolor="#FFFFFF" height="1">
                  <div align="left"><marquee scrollamount=2><%=rs1("message")%></marquee></div>
                </TD>
                <TD width="12%" bordercolor="#FFFFFF" height="1">
                  <div align="center"><img src=<%=rs1("liwu")%>></div>
                </TD>
         	<td valign=middle class=tablebody1 baseName=images/img_visual/blank.gif width="84" height="84" border="0" itemgender="" itemno=<%=rs1("itemid")%> layerno="" localpic=<%=LocalPic%> ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>

		<%if lcase(trim(rs1("tousername")))=lcase(trim(checkStr(request.cookies("aspsky")("username")))) then%>
		<TD width="27%" bordercolor="#FFFFFF" height="1"><a href=z_marry_agree.asp?agree=yes&id=<%=rs1("id")%>&name=<%=rs1("username")%>>同意</a><br><a href=z_marry_agree.asp?agree=no&id=<%=rs1("id")%>&name=<%=rs1("username")%>>不同意</a></TD>
		<%end if%>
					
				  <% rs1.movenext
			  loop
			  rs1.close
			  set rs1=nothing 
			  end if %>
             -->
				<%
				dim rsl
				set rs1=server.createobject("adodb.recordset")
				rs1.open"select * from qiuhun where dlg=false order by addtime desc",connj,3,1
				if rs1.bof then
					response.write "<tr><td>还没有任何人求婚呢</td></tr>"
				else
					do while not rs1.eof
						%>
						<tr>
							<td><b>求婚者</b>：<a href=dispuser.asp?name=<%=rs1("username")%> target="_blank" title="查看<%=rs1("username")%>的资料"><%=rs1("username")%></a></td>
							<td valign=middle rowspan=5 class=tablebody1 baseName=images/img_visual/blank.gif width="84" height="84" border="0" itemgender="" itemno=<%=rs1("itemid")%> layerno="" localpic=<%=LocalPic%> ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
							<!--<td rowspan=4><marquee scrollamount=2 direction=up><%=rs1("message")%></marquee></td>-->
						</tr>
						<tr><td><b>意中人</b>：<a href=dispuser.asp?name=<%=rs1("tousername")%> target="_blank" title="查看<%=rs1("tousername")%>的资料"><%=rs1("tousername")%></a></td></tr>
						<tr><td><font color=blue><marquee scrollamount=2><%=rs1("message")%></marquee></font></td></tr>
						<tr><td><b>求婚时间</b>：<%=rs1("addtime")%></td></tr>
						<tr><td><b>操作</b>：
							<%if lcase(trim(rs1("tousername")))=lcase(trim(checkStr(request.cookies("aspsky")("username")))) then%>
								<a href=z_marry_agree.asp?agree=yes&id=<%=rs1("id")%>&name=<%=rs1("username")%>><font color=white>同意</font></a> | <a href=z_marry_agree.asp?agree=no&id=<%=rs1("id")%>&name=<%=rs1("username")%>><font color=white>不同意</font></a></td></tr>
							<%end if
						rs1.movenext
						%><tr><td colspan=2>&nbsp;</td></tr><%
					loop
				end if%>
			 </TBODY></TABLE></center>
            </div>
          </TD></TR></TABLE>
      <p>
      <BR><BR></TD>
    <TD bgColor=#FFFFFF width=244 height="360" background="images/img_marry/06.jpg">
      <TABLE border=0 cellPadding=10 cellSpacing=15 class=ct-def2 width=185>
        <TR bgColor=#ffffff>
          <TD height=150 bgcolor="#FF7D6B"><IMG height=1 src="images/img_marry/empty.gif" 
            width=1> 
<DIV style="HEIGHT: 50px; WIDTH: 150px">
            <MARQUEE direction=up height=150 onmouseout=this.start() 
            onmouseover=this.stop() scrollAmount=1 scrollDelay=30>在申请注册登记之前，请先了解<a href=z_marry_statute.asp><FONT color=#cc0000>论坛的婚姻法。</FONT></a><BR>您可以去商店的购买<a href=z_visual.asp?shopid=209><FONT color=#cc0000>定情信物</FONT></a>，一定要选一个心上人喜欢的礼物哦。<BR>如果您的情人已经申请向您求婚，请被求婚人点击<a href=z_marry_manager.asp><FONT color=#cc0000>“月老版块”</FONT></a>来完成这一生的承诺。<BR>若想参加婚礼，请点击<a href=z_marry_cabaret.asp><FONT color=#cc0000>“社区酒店”</FONT></a>。</MARQUEE></DIV></TD></TR></TABLE></TD>
    <TD bgColor=#FFFFFF width=4 rowspan="2" height="1">　</TD></TR>
  <tr>
    <TD bgColor=#FFFFFF vAlign=top width=295 height="32" background="images/img_marry/05.jpg">　</TD>
    <TD bgColor=#FFFFFF width=244 height="90" valign="top" background="images/img_marry/07.jpg">
      　</TD>
    </tr>
  </TBODY></TABLE>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
  <TBODY> 
  <TR>
    <TD bgColor=#FFFFFF>　</TD></TR></TBODY></TABLE>
	<% 
	end sub %>

