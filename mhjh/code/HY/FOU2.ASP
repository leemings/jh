<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 会员,体力,内力,攻击,防御 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.close
response.end
else
hy=rs("会员")
tl=rs("体力")
nl=rs("内力")
gj=rs("攻击")
fy=rs("防御")
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
else
if hy=false then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href = "javascript:history.back()"
</script>
<%
else
jiu=request.form("jiu")
	select case jiu
           case "lao"
              tili=40
              yin=2000
           case "wu"
              tili=60
              yin=3000
           case "du"
              tili=80
              yin=4000
           case "mao"
              tili=100
              yin=5000
          end select
	sql="select 银两,道德 from 用户 where 姓名='" & username & "' and 状态<>'死亡' and 状态<>'入狱' and 状态<>'逮捕'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		mess=username & "，你不能操作！"
	else
		if rs("银两")<yin then
			mess=username & "，你的银两不够！"
                    else
                daode=rs("道德")
                jiuliang=int((daode)/100)
                if jiuliang>1 then
                       	sql="update 用户 set 银两=银两-" & yin & ",道德=道德+"& tili &" where 姓名='" & username & "'"
			conn.execute sql
			mess=username & "从相国寺出来后，觉得自己得益非浅，同时也悟出不少道理!"
                else	
                       shi=0.0416*1
			sql="update 用户 set 银两=银两-" & yin & ",道德=道德+"& tili &",最后登录时间=now()+" & shi & ",状态='眠' where 姓名='" & username & "'"
			conn.execute sql
			mess=username & "，由于你和寺内助持谈的投机，于是留你在寺内吃了饭再走，请在1小时后使用该帐号！"
            session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../fou.asp  target=_top"
</script>
<%
		end if
	end if
  end if
%>

<head>
<link rel="stylesheet" href="../style.css" type="text/css">
</head>
<body background='../chatroom/bg1.gif' topmargin='100'>
<table border="1" bgcolor="#00A200" align="center" width="350" cellpadding="10"
cellspacing="13">
  <tr>
    <td bgcolor="#005b00">
      <table width="100%">
        <tr>
          <td>
            <p align="center" style="font-size:14;color:red"><%=mess%></p>
            <p align="center"><a href="fou.asp">返回</a></p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
<%
end if
end if
end if
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing   
%>