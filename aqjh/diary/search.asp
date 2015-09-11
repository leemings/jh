<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
%>
<html><head><title>日记查询结果</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<body bgcolor="#FFFFFF" text="#000000">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;"></DIV>
<SCRIPT LANGUAGE="JavaScript" src="images/overlib.js"></SCRIPT>
<table align=center border=0 cellpadding=0 cellspacing=0 width=574>
<tr><td><img height=13 src="images/top.gif" width=574></td></tr>
<tr><td align=right background=images/top1.gif><table border=0 cellpadding=0 cellspacing=0 width="94%">
<tr><td align=right><table border=0 cellpadding=0 cellspacing=0 width="80%">
<tr align=right><td><a href="mydiary.asp"><img alt=进入我的日记本 border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=最新日记 border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
<td><a href="top.asp?fs=2"><img alt=日记排行 border=0 height=21 src="images/anniu4.gif" width=78></a></td></tr>
</table></td><td rowspan=3 valign=bottom width=26><img height=289 src="images/right.gif" width=8></td>
</tr><%If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
sub diaryclose(mess)
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('提示："& mess &"');location.href = 'list.asp';</script>"
	response.end
end sub
search=cstr(server.HTMLEncode(Request.QueryString("search")))
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记 where (日记标题 like '%"&search&"%'or 日记内容 like '%"&search&"%') Order By 书写日期 DESC", conn1, 1,1
if rs1.eof and rs1.bof then
%><tr><td align=center >对不起，没有找到任何记录!</td></tr><%
else
rs1.PageSize=15'设置每页记录数,可修改          
Dim TotalPages              
TotalPages = rs1.PageCount              
If CurPage>rs1.Pagecount Then               
	CurPage=rs1.Pagecount              
end if                            
rs1.AbsolutePage=CurPage
rs1.CacheSize = rs1.PageSize'设置最大记录数  
Dim Totalcount
Totalcount =INT(rs1.recordcount)              
StartPageNum=1              
do while StartPageNum+10<=CurPage              
StartPageNum=StartPageNum+10              
Loop              
EndPageNum=StartPageNum+9              
If EndPageNum>rs1.Pagecount then EndPageNum=rs1.Pagecount
%><tr valign="top"><td>
 <table align=center bgcolor=#e0d2c5 border=3 cellpadding=2 cellspacing=0 width="98%" >
<%if Request.QueryString("fs")<>"" then%>
<tr align=right bgcolor=#ede5dd>
 <td colSpan=2 align="center">[<%=name%>]创建于：<font color=blue><%=year(bb8)%>年<%=month(bb8)%>月<%=day(bb8)%>日</font> 总数：<font color=blue><%=bb6%></font> 鲜花：<font color=blue><%=bb4%></font>
鸡蛋：<font color=blue><%=bb5%></font> 留言：<font color=blue><%=bb7%></font> 文采:<font color=blue><%=bb9%></font><br>
『<a href="list.asp?fs=1&name=<%=rs1("用户名")%>&zz=1" ><font color=blue><%=mylb(1)%></font></a> 』『<a href="list.asp?fs=1&name=<%=rs1("用户名")%>&zz=2" ><font color=blue><%=mylb(2)%></font></a> 』『<a href="list.asp?fs=2&name=<%=rs1("用户名")%>&zz=3" ><font color=blue><%=mylb(3)%>』</font></a></td>
</tr>
<%end if
I=0
p=rs1.PageSize*(Curpage-1)              
do while (Not rs1.Eof) and (I<rs1.PageSize)              
p=p+1
select case rs1("保密")
case "1"
	bm="公开欣赏"
case "2"
	bm="保密内容!"
case "3"
	bm="凭密码查看"
case "4"
	bm="["&rs1("保密条件") &"]查看"
case "5"
	bm="["&rs1("保密条件")&"]级以上"
case "6"
	bm="["&rs1("保密条件")&"]查看"
end select
if i/2=int(i/2) then
%>
              <tr align=middle bgcolor='cccccc' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='cccccc'> 
                <td width="63%" > 
                  <div align=left><img height=14 src="images/hot.gif"width=11><a href="show.asp?id=<%=rs1("id")%>&wj=list.asp" onMouseOver="drs('<div align=center><b>∷信息资料∷</b></div>鲜花:<font color=#ff0000><%=rs1("鲜花")%></font>朵<br>鸡蛋:<font color=#ff0000><%=rs1("鸡蛋")%></font>个<br>留言:<font color=#ff0000><%=rs1("留言数")%></font>条<br>方式:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("日记标题")%></a></div>
                </td>
                <td width="37%"> 
                  <div align=left><a href="list.asp?fs=1&name=<%=rs1("用户名")%>" target="_blank"><img src="images/gd.gif" width="13" height="13" border="0" alt="查看作者专辑！"></a>&nbsp;作者:<a href="mydiary.asp?name=<%=rs1("用户名")%>" title="查看作者日记本" target="_blank"><%=rs1("用户名")%></a>&nbsp;<%=rs1("书写日期")%></div>
                </td>
              </tr>
              <%else%>
              <tr align=middle bgcolor='efefef' onmouseover=this.bgColor='#00B0DD'; onmouseout=this.bgColor='efefef'> 
                <td width="63%" > 
                  <div align=left><img height=14 src="images/hot.gif"width=11><a href="show.asp?id=<%=rs1("id")%>&wj=list.asp" onMouseOver="drs('<div align=center><b>∷信息资料∷</b></div>鲜花:<font color=#ff0000><%=rs1("鲜花")%></font>朵<br>鸡蛋:<font color=#ff0000><%=rs1("鸡蛋")%></font>个<br>留言:<font color=#ff0000><%=rs1("留言数")%></font>条<br>方式:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("日记标题")%></a></div>
                </td>
                <td width="37%"> 
                  <div align=left><a href="list.asp?fs=1&name=<%=rs1("用户名")%>" target="_blank"><img src="images/gd.gif" width="13" height="13" border="0" alt="查看作者专辑！"></a>&nbsp;作者:<a href="mydiary.asp?name=<%=rs1("用户名")%>" title="查看作者日记本" target="_blank"><%=rs1("用户名")%></a>&nbsp;<%=rs1("书写日期")%></div>
                </td>
              </tr>
              <%end if
              I=I+1
rs1.MoveNext              
Loop
%>
              <form method="GET" action="search.asp">
                <tr align=right bgcolor=#ede5dd> 
                  <td colSpan=2 align="center">关键字： 
                    <input type="text" name="search" size=10 onfocus=this.select() onmouseover=this.focus()>
                    <input name="B1" type="submit" value="搜索" >
                    共<%=Totalcount%>篇日记[ 
                    <% if StartPageNum>1 then %>
                    <a href="list.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
                    <%end if%>
                    <% For I=StartPageNum to EndPageNum
					if I<>CurPage then %>
                    <a href="list.asp?CurPage=<%=I%>"><%=I%></a> 
                    <% else %>
                    <b><font color=#ff0000><%=I%></font></b> 
                    <% end if %>
                    <% Next %>
                    <% if EndPageNum<rs1.Pagecount then %>
                    <a href="list.asp?CurPage=<%=EndPageNum+1%>">>></a> 
                    <%end if%>
                    ] </td>
                </tr>
              </form>
            </table>
</td>
</tr>
<%end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing%>
</table>
</td></tr><tr><td><img height=43 src="images/down.gif" width=574></td></tr></table></body></html>
