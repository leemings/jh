<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
sql="SELECT 门派,体力 FROM 用户 where 姓名='" & aqjh_name &"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "你不是江湖中人"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
else
	pai=rs("门派")
	tl=rs("体力")
rs.close
sql="SELECT q,s FROM p where a='"&pai&"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	shengming="没有门派战斗力"
	shidi="没有敌人"
	else
	shengming=rs("s")
	shidi=rs("q")
	end if
rs.close
%>
<html><head>
<meta http-equiv="Content-Language" content="zh-tw">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>门派大战</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../BG.gif" bgproperties="fixed">
<div align="right">
  <table border="1" width="100%" cellspacing="0" cellpadding="2" bgcolor="#008000" bordercolordark="#FFFFFF" bordercolorlight="#000000" background="../BG.gif" align="right">
    <tr>
      <td width="100%" align="center"><img border="0" src="../img/ki17.gif" width="18" height="18">【门派大战】<img border="0" src="../img/ki17.gif"></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao.asp">申请保护</a>:<font color="#FFFFCC">申请门派保护</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao2.asp">关闭保护</a>:<font color="#FFFFCC">关闭门派保护</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/门派挑战$')">门派挑战</a>:<font color="#FFFFCC">向他门派挑战</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/门派求和$')">门派求和</a>:<font color="#FFFFCC">向他门派求饶</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mp2.asp">增加战斗</a>:<font color="#FFFFCC">增加门派战斗</font></td>
    </tr>
    <tr>
      <td width="100%">个人战斗:<font color="#FFFFCC" size="2"><%=tl%></font></td>
    </tr>
    <tr>
      <td width="100%">门派战斗:<font color="#FFFFCC" size="2"><%=shengming%></font></td>
    </tr>
    <tr>
      <td width="100%">世仇:<input type="text" size="12" value="<%=shidi%>" readonly></td>
    </tr>
    <tr>
      <td width="100%" align="center"> <table cellpadding="3" cellspacing="0" border="0" align="center" width="100%">
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center">招式</div>
      </td>
      <td width="29%"> 
        <div align="center">武功</font></div>
      </td>
      <td width="28%"> 
        <div align="center">内力</font></div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM y where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"> <font color="#ff0000" size="2"> <a href="javascript:s('/门派大战$ <%=rs("a")%>')" title="此招试为：<%=dj%>重!"><%=rs("a")%></a></font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("c")%></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("d")%></font></div>
      </td>
    </tr>
 <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if%>
</table>
</td>
    </tr>
    <tr><td width="100%" align="center">『爱情江湖总站』</td>
    </tr>
  </table></div></body></html>