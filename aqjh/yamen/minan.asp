<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from l where d='人命' and a<now()-10",conn,3,3
seekname=Trim(Request.Form("name"))
if seekname<>"" then
	rs.open "Select * From l where d='人命' and (b='"&seekname&"' or c='"&seekname&"') Order by a DESC",conn,3,3
else
	rs.open "Select * From l where d='人命' Order by a DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title><%=Application("aqjh_chatroomname")%>-江湖命案</title>
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<br>
<br>
<table width="778" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="482" valign="top" background="../images/man1.gif"><img src="../images/ma3.gif" width="143" height="36"></td>
    <td width="163" background="../images/man1.gif"><img src="../images/ma2.gif" width="194" height="80"></td>
    <td valign="bottom" align="right" width="133" background="../images/man1.gif"><img src="../images/ma4.gif" width="69" height="22"></td>
  </tr>
  <tr bgcolor="#000000"> 
    <td colspan="3" valign="top" height="7"></td>
  </tr>
</table>
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
  <tbody> 
  <tr> 
    <td height="38" width="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
        <tr> 
          <td width="64" height="3"> &nbsp; 
            <%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %>
          </td>
          <td width="223" height="3">[ <font color="#FF6600">江 湖 命 案 </font>] 
          </td>
          <td width="261" height="3"> 
            <form method="POST" action="minan.asp" name="login">
              <div align="center">输入查询姓名回车： 
                <input type="text" name="name" value="<%=aqjh_name%>"size="15" maxlength="10" style="border: 1px solid; font-size: 9pt; border-color:#000000 solid">
              </div>
            </form>
          </td>
          <td width="175" height="3"> <%=Application("aqjh_chatroomname")%><font color="#000000">程序修改</font> 
          </td>
        </tr>
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<div align="center"> 
  <table width="91%" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr> 
      <td width="100%"> 
        <table border="1" cellspacing="1" cellpadding="0" width="97%"
bordercolorlight="#EFEFEF" bordercolor="#000000" bgcolor="#eeeeee">
          <tr> 
            <td width="20%"> 
              <div align="center"><font color="#FF6600"> 死 亡 者 </font></div>
            </td>
            <td width="25%"> 
              <div align="center"><font color="#FF6600"> 时 间 </font></div>
            </td>
            <td width="19%"> 
              <div align="center"><font color="#FF6600"> 杀 人 凶 手 </font></div>
            </td>
            <td width="36%"> 
              <div align="center"><font color="#FF6600"> 死 亡 原 因 </font></div>
            </td>
          </tr>
          <%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
       <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
            <td width="20%"> 
              <div align="center"> <font color="#FF3333"><%=rs("b")%></font> 
              </div>
            </td>
            <td width="25%"> 
              <div align="center"> <%=rs("a")%> </div>
            </td>
            <td width="19%"> 
              <div align="center"> <%=rs("c")%> </div>
            </td>
            <td width="36%"> 
              <div align="center"> <%=rs("e")%> </div>
            </td>
          </tr>
          <%rs.movenext%>
          <%count=count+1%>
          <%loop
			pa=rs.pagecount
			mu=musers()
          rs.close
          set rs=nothing
          conn.close
          set conn=nothing%>
        </table>
        <table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
          <tr> 
            <td align="left" width="37%">[共<font color="red"><b><%=pa%></b></font>页][<font
color="red"><b><%=mu%></b></font>人死亡]</td>
            <td align="right" width="63%"> 
              <div align="center">[<a href="minan.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="minan.asp?page=<%=page+1%>">下一页</a>]</div>
            </td>
          </tr>
        </table>
  </table>
</div>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 杀人 from l where d='人命'")
musers=tmprs("杀人")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body>