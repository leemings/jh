<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人？！');window.close();</script>"
	response.end
end if
sjjh_id=rs("id")
rs.close
conn.close
set conn=nothing
function chuser(u)
	dim filter,xx,usernameenable,su
	for i=1 to len(u)
		su=mid(u,i,1)
		xx=asc(su)
		zhengchu = -1*xx \ 256
		yushu = -1*xx mod 256
		if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
			chuser=true
			exit function
		end if
	next
	chuser=false
end function
wt=lcase(trim(request.form("newwt")))
da=lcase(trim(request.form("newda")))
oldwt=lcase(trim(request("oldwt")))
oldda=lcase(trim(request("oldda")))
if wt="" then
	Response.Write "<script Language=Javascript>alert('错误：提示问题不可为空！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if da="" then
	Response.Write "<script Language=Javascript>alert('错误：答案不可为空！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if len(wt)>5 then
	Response.Write "<script Language=Javascript>alert('错误：提示问题最多为五个汉字！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if len(da)>5 then
	Response.Write "<script Language=Javascript>alert('错误：问题答案最多为五个汉字！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if chuser(wt) or chuser(oldwt)then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
if chuser(da) or chuser(oldda) then 
	Response.Redirect "../error.asp?id=60"
	response.end
end if
if instr(wt,"or")<>0 or instr(da,"or")<>0 or instr(wt,"=")<>0 or instr(da,"=")<>0 or instr(oldwt,"or")<>0 or instr(oldda,"or")<>0 or instr(oldwt,"=")<>0 or instr(oldda,"=")<>0then
	Response.Redirect "../error.asp?id=54"
	response.end
end if
if trim(request.form("newwt"))="" or trim(request.form("newda"))="" then
	Response.Write "<script Language=Javascript>alert('错误：提示问题和答案都不可为空！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if wt=da then
	Response.Write "<script Language=Javascript>alert('错误：提示问题和答案不可以相同！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if left(wt,3)="%20" OR InStr(wt,"=")<>0 or InStr(wt,"`")<>0 or InStr(wt,"'")<>0 or InStr(wt," ")<>0 or InStr(wt,"　")<>0 or InStr(wt,"'")<>0 or InStr(wt,chr(34))<>0 or InStr(wt,"\")<>0 or InStr(wt,",")<>0 or InStr(wt,"<")<>0 or InStr(wt,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
if left(oldwt,3)="%20" OR InStr(oldwt,"=")<>0 or InStr(oldwt,"`")<>0 or InStr(oldwt,"'")<>0 or InStr(oldwt," ")<>0 or InStr(oldwt,"　")<>0 or InStr(oldwt,"'")<>0 or InStr(oldwt,chr(34))<>0 or InStr(oldwt,"\")<>0 or InStr(oldwt,",")<>0 or InStr(oldwt,"<")<>0 or InStr(oldwt,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
if left(da,3)="%20" OR InStr(da,"=")<>0 or InStr(da,"`")<>0 or InStr(da,"'")<>0 or InStr(da," ")<>0 or InStr(da,"　")<>0 or InStr(da,"'")<>0 or InStr(da,chr(34))<>0 or InStr(da,"\")<>0 or InStr(da,",")<>0 or InStr(da,"<")<>0 or InStr(da,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
if left(oldda,3)="%20" OR InStr(oldda,"=")<>0 or InStr(oldda,"`")<>0 or InStr(oldda,"'")<>0 or InStr(oldda," ")<>0 or InStr(oldda,"　")<>0 or InStr(oldda,"'")<>0 or InStr(oldda,chr(34))<>0 or InStr(oldda,"\")<>0 or InStr(oldda,",")<>0 or InStr(oldda,"<")<>0 or InStr(oldda,">")<>0 then
	Response.Redirect "../error.asp?id=120"
	response.end
end if
%><!--#include file="data.asp"--><%
sql="select * from 密码保护 where 用户="&sjjh_id
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('你还未申请过密码保护！');window.close();</script>"
	response.end
end if
a=trim(rs("问题"))
b=trim(rs("答案"))
if oldwt<>a or oldda<>b then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('滚，你想做什么！');window.close();</script>"
	response.end
end if
rs.close
sql="update 密码保护 set 问题='" & wt & "',答案='" & da & "' where 用户="&sjjh_id
set rs=conn1.execute(sql)
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>密码保护资料修改完成</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER219.JPG">
<div align="center"> 
  <form name="form1" method="post" action="">
    <p><font size="+1" color="#0000FF">您的密码保护资料修改成功</font></p>
    <p>请您一定要牢记自已的<b><font color="#FF0000">ID</font></b>和密码<b><font color="#FF0000">提示问题</font></b>及<font color="#FF0000"><b>答案</b></font>，这是您找回密码的唯一办法</p>
    <p><font color="#0000FF">您的ID为：</font><font color="#FF0000"><%=sjjh_id%></font></p>
    <p><font color="#0000FF">用户名为：</font><font color="#FF0000"><%=sjjh_name%></font></p>
    <p><font color="#0000FF">新的提示问题为：</font><font color="#FF0000"><%=wt%></font></p>
    <p><font color="#0000FF">新的问题答案为：</font><font color="#FF0000"><%=da%></font></p>
    <p>
	  <input type=button value=关闭窗口 onClick='window.close()' name="button">
    </p>
  </form>
  <p>　</p>
  </div>
<p align="center"><font color="#FF0000" size="2"><b>『E线江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:一线天</b></font></p>
</body>
</html>