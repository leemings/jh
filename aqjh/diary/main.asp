<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if request("name")="" then
name=aqjh_name
else
name=request("name")
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记用户 where 姓名='"& name &"'", conn1, 1,1
if rs1.eof and rs1.bof then
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
Response.Redirect "reg.asp"
end if
rjname=rs1("日记本名")
lys=rs1("留言数")
jds=rs1("鸡蛋")
xhs=rs1("鲜花")
sm=rs1("说明")
wcs=rs1("文采")
dim mylb(3)
mylb(1)=rs1("lb1")
mylb(2)=rs1("lb2")
mylb(3)=rs1("lb3")
rjs=rs1("日记数")
jl=rs1("建立日期")
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html><head><title><%=name%>的日记本</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css><body bgcolor="#ebe4de" text="#FFFFFF">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;"></DIV>
<SCRIPT LANGUAGE="JavaScript" src="images/overlib.js"></SCRIPT>
<table align=center border=0 cellpadding=0 cellspacing=0 width=574>
<tr><td><img height=14 src="images/top-1.gif" width=574></td></tr><tr>
    <td align=middle background=images/top1-1.gif height="2">
      <table border=0 cellpadding=2 cellspacing=0 width="100%">
        <tr> 
          <td align=middle width="1%">&nbsp;</td>
          <td align=middle width="20%"> 『<font color=blue><%=rjname%></font>』</td>
          <td width="11%" rowspan="7">&nbsp;</td>
          <td valign=top width="63%" rowspan="7"> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tr align=right> 
                <td><a href="list.asp"><img alt=最新日记 border=0 height=21 src="images/anniu2.gif" width=80></a></td>
                <td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
                <td><a href="top.asp?fs=2"><img alt=日记排行 border=0 height=21 src="images/anniu4.gif" width=78></a></td>
              </tr>
            </table>
            <div align="center">『<a href="main.asp?fs=1&name=<%=name%>" ><font color=blue>全部</font></a>』『<a href="main.asp?fs=2&name=<%=name%>&zz=1" ><font color=blue><%=mylb(1)%></font></a> 
              』『<a href="main.asp?fs=2&name=<%=name%>&zz=2" ><font color=blue><%=mylb(2)%></font></a> 
              』『<a href="main.asp?fs=2&name=<%=name%>&zz=3" ><font color=blue><%=mylb(3)%></font></a>』</div>
            <br>
            <table border=1 cellpadding=2 cellspacing=0 width="95%" align="right">
              <%If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
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
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
select case Request.QueryString("fs")
case "1"
	name=trim(Request.QueryString("name"))
	rs1.Open "Select * From 日记用户 where 姓名='"& name &"'", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("此用户资料有问题，无法查看！")
	end if
	bb4=rs1("鲜花")
	bb5=rs1("鸡蛋")
	bb6=rs1("日记数")
	bb7=rs1("留言数")
	bb8=rs1("建立日期")
	bb9=rs1("文采")
	mylb(1)=rs1("lb1")
	mylb(2)=rs1("lb2")
	mylb(3)=rs1("lb3")
	rs1.close
	rs1.Open "Select * From 日记 where 用户名='"& name &"' Order By 书写日期 DESC", conn1, 1,1
case "2"
	name=trim(Request.QueryString("name"))
	lb=trim(Request.QueryString("zz"))
	rs1.Open "Select * From 日记用户 where 姓名='"& name &"'", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("此用户资料有问题，无法查看！")
	end if
	bb4=rs1("鲜花")
	bb5=rs1("鸡蛋")
	bb6=rs1("日记数")
	bb7=rs1("留言数")
	bb8=rs1("建立日期")
	bb9=rs1("文采")
	mylb(1)=rs1("lb1")
	mylb(2)=rs1("lb2")
	mylb(3)=rs1("lb3")
	rs1.close
	rs1.Open "Select * From 日记 where 用户名='"& name &"' and lb="&lb&" Order By 书写日期 DESC", conn1, 1,1
case else
	rs1.Open "Select * From 日记 where 用户名='"& name &"' Order By 书写日期 DESC", conn1, 1,1
end select
if rs1.eof and rs1.bof then
%><tr><td align=center >没有记录</td></tr></table>
<%
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
%>
              <tr> 
                <td align=left ><img height=14 src="images/hot.gif" width=11><a href="show.asp?id=<%=rs1("id")%>&wj=main.asp" onMouseOver="drs('<div align=center><b>∷信息资料∷</b></div>鲜花:<font color=#ff0000><%=rs1("鲜花")%></font>朵<br>鸡蛋:<font color=#ff0000><%=rs1("鸡蛋")%></font>个<br>留言:<font color=#ff0000><%=rs1("留言数")%></font>条<br>方式:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("日记标题")%></a></td>
                <td height=20 width="24%">(<%=rs1("书写日期")%>)</td>
              </tr>
              <%
I=I+1
rs1.MoveNext              
Loop
%>
              <tr> 
                <td colspan="2"> 
                  <div align=right> 共<%=Totalcount%>篇[ 
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
                    ] </div>
                </td>
              </tr>
            </table>
            <%end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing%>
          </td>
          <td valign=bottom width="5%" rowspan="7"><img src="images/right.gif" width="8" height="289" align="bottom"></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> 总数：<font color=blue><%=rjs%></font>&nbsp;&nbsp;留言：<font color=blue><%=lys%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> 鲜花：<font color=blue><%=xhs%></font>&nbsp;&nbsp;鸡蛋：<font color=blue><%=jds%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> 文采：<font color=blue><%=wcs%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%" valign="top"><a href="add.asp"><img alt=写江湖日记 border=0 height=24 src="images/xie.gif" width=65></a><br>
            <a href="reg.asp?action=edit"><img alt=修改资料 border=0 height=24 src="images/xiugai.gif" width=76></a> 
          </td>
        </tr>
        <tr> 
          <td align=left colspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td align=left colspan="2" valign="top" height="2"> 
            <div align="left"><br>
            </div>
          </td>
        </tr>
      </table>
    </td>
</tr><tr>
    <td align="right"><img height=44 src="images/down-1.gif" width=574></td>
  </tr></table>
<br>
</body></html>
