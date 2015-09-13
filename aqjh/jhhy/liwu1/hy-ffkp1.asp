<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../pass.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
function chuser(u)
	dim filter,xx,usernameenable,su
	for i=1 to len(u)
		su=mid(u,i,1)
		xx=asc(su)
		zhengchu = -1*xx \ 256
		yusu = -1*xx mod 256
		if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yusu=129 or yusu>192 or (yusu<2 and yusu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yusu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
			chuser=true
			exit function
		end if
	next
	chuser=false
end function
name=LCase(trim(Request.form("name")))
psw=trim(Request.form("psw"))
rz=trim(Request.Form("psw2"))
rzjm=trim(request.form("regjm"))
nowinroom=session("nowinroom")
for each element in Request.Form
	if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('江湖提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
next
name=trim(name)
name=server.HTMLEncode(name)
if name="无" or name="未定" or name="" or name=null then response.redirect "error.asp?id=130"
if chuser(name) then Response.Redirect "error.asp?id=120"
if len(psw)<5 then Response.Redirect "error.asp?id=52"
for i=1 to len(name)
	usernamechr=mid(name,i,1)
	if asc(usernamechr)>0 then  Response.Redirect "error.asp?id=120"
next
if instr(name,"or")<>0 or instr(psw,"or")<>0 then Response.Redirect "error.asp?id=54"
if instr(name,"=")<>0 or instr(psw,"=")<>0 then Response.Redirect "error.asp?id=54"
if rz<>rzjm then Response.Redirect "error.asp?id=166"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name,"　")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 or instr(name,chr(13)) then Response.Redirect "error.asp?id=120"
psw=md5(psw)
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
if sj<>"2015-07-01" then
	Response.Write "<script Language=Javascript>alert('对不起，节日期间才发放！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,w5,等级,会员等级 FROM 用户 WHERE 姓名='" & name & "' and 密码='"&psw&"'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你现在还不是我们江湖的人吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("等级")<=40 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('为防止作弊，只有等级在40级以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
hy=rs("会员等级")
if hy<2 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你不是二级会员以上付费会员，请回吧!');history.go(-1)';}</script>"
		response.end
end if
myid=rs("id")
wpdate=rs("w5")
hydj=rs("会员等级")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.asp")
rs1.open "select * from a where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已经领过卡片了！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs1.close
lw=""
randomize timer
r=int(rnd*2)+1
select case r
	case 1 '福神卡
		n=hydj
		temp=add(wpdate,"福神卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="福神卡"&n&"个"
	case 2 '涨钱卡
		n=hydj
		temp=add(wpdate,"涨钱卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myid
		lw="涨钱卡"&n&"个"
       case 3 '健身卡
		n=hydj
		temp=add(wpdate,"健身卡",n)
		sql="update 用户 set w5='"&temp&"' where id="&myide
		lw="健身卡"&n&"个"	
	
end select
set rs=conn.execute(sql)
conn1.execute "insert into a(yhid,姓名,礼物) values ('"&myid&"','"&name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
conn1.close
set conn1=nothing
says="<font size=2 color=red>【领取节日礼物】"&name&"</font><font size=2 color=blue>是<font color=red>"&hydj&"</font>级会员，所以去节日礼物领取了<font color=red>"&lw&"</font>，会员等级越高领得卡片越多!</font>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<title>节日礼品发放组委会</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FF9900" background="bg.gif">
<table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC">
  <tr bgcolor="#FFFFFF" align="center"> 
      <td width="367" height="252" bgcolor="#FF9900" background="bg.gif"><div align="center"><b>恭喜你从<font color="#FF0000">快乐江湖</font>总站那里得到了<br>
        <font color="#66FF00"><%=lw%></font></b><br>
<br>
<br>
<br>
<br>
[快乐江湖网]</div></td>
  </tr>
</table>
</body>
</html>
