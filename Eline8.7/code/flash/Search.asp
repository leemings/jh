<!--#include file="conn.asp"-->
<!--#include file="func.asp"-->
<!--#include file="function.asp"-->
<%
keyword=trim(request("keyword")) 
keyword=replace(keyword,"'","''") 
stype=request("stype")
if keyword="" then
errmsg=errmsg+"查找字符不能为空，请重输入查找的信息<a href=""javascript:history.go(-1)"">返回重查</a>"
call error()
Response.End 
end if
%>
<html>
<head>
<title>★E线江湖总站-Flash频道☆</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:eline_email@etang.com;QQ:88617427">
<link rel="stylesheet" href="index2.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" link="#000000" vlink="#000000" alink="#000000">
<!--#include file="top.asp"-->
<CENTER>
  <TABLE border="0" width="777" cellpadding="0" cellspacing="0">
    <TBODY>
    <TR>
        <TD background="list_pic/image1.jpg" width="160" align="center" valign="top"><table width="197" border="0" cellspacing="0" cellpadding="0" height="707">
            <tr> 
              <td bgcolor="#FFFFF0" align="center" width="195" height="30"><font color=blue size="3">↓最 新 上 传↓</font></td>
            </tr>
            <tr> 
              <td bgcolor="#FFFFF0" width="195"> 
                <%
	dim rs
	dim sql
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql ="Select top 10 * From flash Order By ID DESC"
	rs.open sql,Conn,1,1
	do while not rs.eof
%>
                <span class=b>·</span><a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" ><%=rs("title")%></a> 
                <span class="s">(<%=month(rs("time"))%>-<%=day(rs("time"))%>)</span> 
                <br> 
                <%  
	rs.MoveNext 
	Loop 
	rs.close 
%>
              </td>
            </tr>
            <tr> 
              <td height="30" bgcolor="#FFFFF0" align="center" width="195"><font color=blue size="3">↓人 气 TOP10↓</font></td>
            </tr>
            <tr> 
              <td bgcolor="#FFFFF0" width="195"> 
                <% 
dim rs1,sql1 
set rs1=server.createobject("adodb.recordset") 
sql1="select * from flash ORDER BY hits desc" 
rs1.open sql1,conn,1,1 
if rs1.eof and rs1.bof then  
response.write "<p align='center'>没有热点动画</p>" 
else  
do while not rs1.eof 
dim k 
%>
                <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
                [<span class="s"><font color="#FF0000"><%=rs1("hits")%></font></span>]<br> 
                <%k=k+1 
	if k=10 then exit do 
	rs1.movenext 
	loop 
	end if 
	rs1.close 
%>
              </td>
            </tr>
            <tr> 
              <td height="18" bgcolor="#FFFFF0" width="195"></td>
            </tr>
            <tr> 
              <td bgcolor="#FFFFF0" align="center" width="195" height="30"><font color=blue size="3">↓下 载 排 行↓</font></td>
            </tr>
            <tr> 
              <td bgcolor="#FFFFF0" width="195"> 
                <% 
set rs1=server.createobject("adodb.recordset") 
sql1="select * from flash ORDER BY download desc" 
rs1.open sql1,conn,1,1 
if rs1.eof and rs1.bof then  
response.write "<p align='center'>没有动画下载</p>" 
else  
do while not rs1.eof 
dim n 
%>
                <span class="b">·</span><a href="show.asp?ID=<%=rs1("ID")%>&ClassID=<%=rs1("ClassID")%>"><%=rs1("title")%></a> 
                [<span class="s"><font color="#FF0000"><%=rs1("download")%></font></span>]<br> 
                <%n=n+1 
	if n=10 then exit do 
	rs1.movenext 
	loop 
	end if 
	rs1.close 
	set rs1=nothing 
%>
              </td>
            </tr>
          </table></TD>
      <TD align="center" valign="top" height="261"> 
        <TABLE width="100%" height="707" border="0" cellpadding="0" cellspacing="0">
            <TBODY>
              <TR> 
                <%
mm=4


Set rs=Server.CreateObject("ADODB.Recordset")  
if   stype="name" then
	strsql="select * from flash where Title Like '%"& keyword &"%'"
else
	strsql="select * from flash where Content Like '%"& keyword &"%'"
end if
rs.open strsql,conn,1,1


page=request("page")

rs.PageSize = mm
maxpage=rs.PageCount 
result_num=rs.RecordCount

if result_num<=0 then
	if search="" then
		word="目前还没有记录!"
	else	
		word="没有查到符合条件的记录!"
	end if	
else
	maxpage=rs.PageCount 
	page=request("page")
	
	if Not IsNumeric(page) or page="" then
		page=1
	else
		page=cint(page)
	end if
	
	if page<1 then
		page=1
	elseif  page>maxpage then
		page=maxpage
	end if
	
	rs.AbsolutePage=Page
end if
%>
                <TD width="100%" align="center" valign="top" bgcolor="#FFFFF0"> 
                  <table border="0" width="100%" cellpadding="0" cellspacing="0">
                    <tbody>
                      <tr> 
                        <td valign="bottom" bgcolor="#FFFFF0" height="20"> <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;符合条件的Flash共有<font color="#FF0000"><%=result_num%><font color="#000000">个</font></font>， 
                            <%
prop="keyword="&keyword&"&stype="&stype

maxpage=int(maxpage)
page=int(page)
call ShowLastNext (maxpage,page,prop)
%>
                          </div></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#bfe1ff"></td>
                      </tr>
                    </tbody>
                  </table>
                  <% 
do while not rs.eof 
i=i+1 
%> <table width="546" border="0" cellspacing="0" cellpadding="0" height="134">
                    <tr bgcolor="#FFFFF0"> 
                      <td colspan="2" height="19" width="559" bgcolor="#FFFFF0"> 
                        <hr color="#FFCC66" size="1" width="98%"> </td>
                    </tr>
                    <tr> 
                      <td rowspan="5" width="119" align="center" bgcolor="#FFFFF0" height="115"><a href="show.asp?ID=<%=rs("ID")%>&amp;ClassID=<%=rs("ClassID")%>"><img src="<%=rs("img")%>" width="102" height="102"></a></td>
                      <td height="23" bgcolor="#FFFFF0" width="438"><font color="#666600">动画名称：</font><font color="#FF0000"><%=rs("title")%></font></td>
                    </tr>
                    <tr> 
                      <td height="23" bgcolor="#FFFFF0" width="438"><font color="#666600">上传时间：</font><span class=time><%=rs("time")%></span> 大小：<span class=time><%=rs("KB")%>K</span> <font color="#666600">观看：</font><span class=time><%=rs("hits")%></span> 次 <font color="#666600">下载：</font><span class=time><%=rs("download")%></span> 次 <font color="#666600">评价：</font><%=rs("commend")%></td>
                    </tr>
                    <tr> 
                      <td height="23" bgcolor="#FFFFF0" width="438"><font color="#666600">动画简介：</font> 
                        <% if len (rs("content"))>30 then %> <%= left (rs("content"),28)%>... 
                        <% else %> <%=rs("content")%> <% end if %> </td>
                    </tr>
                    <tr> 
                      <td height="23" bgcolor="#FFFFF0" width="438"><font color="#666600">作者信箱：</font> 
                        <% if rs("email")="" then%>
                        没有该动画作者的信箱 
                        <% else %> <a href="mailto:<%=rs("email")%>"><%=rs("email")%></a> <% end if %> </td>
                    </tr>
                    <tr> 
                      <td height="23" align="center" bgcolor="#FFFFF0" width="438"><img src="images/img_1.gif" width="16" height="16" align="absmiddle"> 
                        <a href="show.asp?ID=<%=rs("ID")%>&ClassID=<%=rs("ClassID")%>" target="_blank">在线观看</a>　　　　<img src="images/download.gif" width="16" height="16" align="absmiddle"> 
                        <a href="download.asp?ID=<%=rs("ID")%>" >下载存盘</a></td>
                    </tr>
                  </table>
                  <%
					rs.MoveNext
					row_count=row_count+1
					 if row_count> rs.PageSize then Exit Do
					Loop
					conn.close
set conn=nothing
set rs=nothing
			 %> <TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
                    <TBODY>
                      <TR> 
                        <TD valign="bottom" bgcolor="#FFFFF0"> <CENTER>
                            符合条件的Flash共有<font color="#FF0000"><%=result_num%><font color="#000000">个</font></font>， 
                            <%
prop="keyword="&keyword&"&stype="&stype

maxpage=int(maxpage)
page=int(page)
call ShowLastNext (maxpage,page,prop)
%>
                          </CENTER></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                </TD>
              </TR>
            </TBODY>
          </TABLE>
        </TD>
    </TR>
  </TBODY>
</TABLE>
</CENTER>
</BODY>
</HTML>
<!--#include file="bottom.asp"-->