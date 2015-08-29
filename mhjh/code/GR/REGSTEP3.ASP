<%@ LANGUAGE=VBScript%>
<!--#include file="pass.asp"-->
<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
    server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
    if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "你提交的路径有误，禁止从站点外部提交数据！"
        response.end
    end if
username=trim(Request.form("username"))
jsr=trim(request.form("jsr"))
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
if chuser(username) then Response.Redirect "../error.asp?id=524"
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") then Response.Redirect "../error.asp?id=056"
next
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
repassword=trim(Request.Form("repassword"))
sex=Request.Form("sex")
country=Request.Form("country")
txxx=trim(Request.Form("txxx"))
jsr=server.HTMLEncode(trim(Request.Form("jsr")))
if Application("yx8_mhjh_systemname")="" then Response.Redirect "../error.asp?id=001"
if account="" or username="" or password="" then Response.Redirect "../error.asp?id=002"
if len(account)<6 or len(username)<2 or len(password)<6 then Response.Redirect "../error.asp?id=004"
if len(account)>15 or len(username)>7 or len(password)>15 then Response.Redirect "../error.asp?id=530"
if len(username)>5 then Response.Redirect "../error.asp?id=245"
if len(country)>1 then Response.Redirect "../error.asp?id=530"
if jsr=username then response.redirect "../error.asp?id=511"
if instr(country,"file:")<>0 or instr(country,"script:")<>0 or instr(country,"'")<>0 or instr(country,chr(34))<>0 or instr(country,"\")<>0 or InStr(country,",")<>0 or InStr(country,"<")<>0 or InStr(country,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(account,"file:")<>0 or instr(account,"script:")<>0 or instr(account,"'")<>0 or instr(account,chr(34))<>0 or instr(account,"\")<>0 or InStr(account,",")<>0 or InStr(account,"<")<>0 or InStr(account,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(username,"file:")<>0 or instr(username,"script:")<>0 or instr(username,"'")<>0 or instr(username,chr(34))<>0 or instr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(username,"未定")<>0 or instr(username,"站长")<>0 or instr(username,"狗")<>0 or instr(username,"系统")<>0 or InStr(username,"网管")<>0 or InStr(username,"爸")<>0 or InStr(username,"妈")<>0 or InStr(username,"操")<>0 or InStr(username,"公告")<>0 or InStr(username,"天高")<>0 or InStr(username,"逼")<>0 or InStr(username,"官府")<>0 or InStr(username,"发钱")<>0 or InStr(username,"赈灾")<>0 or InStr(username,"通知")<>0 or InStr(username,"消息")<>0 then Response.Redirect "../error.asp?id=057"
if instr(password,"file:")<>0 or instr(password,"script:")<>0 or instr(password,"'")<>0 or instr(password,chr(34))<>0 or instr(password,"\")<>0 or InStr(password,",")<>0 or InStr(password,"<")<>0 or InStr(password,">")<>0 then Response.Redirect "../error.asp?id=524"
if InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"　")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=510"
for i=1 to len(account)
accountchr=asc(mid(account,i,1))
if accountchr<48 or (accountchr>57 and accountchr<65) or (accountchr>90 and accountchr<97) then Response.Redirect "../error.asp?id=006"
next
for i=1 to len(username)
usernamechr=mid(username,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=003"
next
if password<>repassword then Response.Redirect "../error.asp?id=007"
for i=1 to len(password)
passwordchr=asc(mid(password,i,1))
if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "../error.asp?id=005"
next
if not(sex="男" or sex="女") then Response.Redirect "../error.asp?id=008"
if not(country="秦" or country="楚" or country="赵" or country="魏" or country="齐" or country="韩" or country="燕") then Response.Redirect "../error.asp?id=244"
illegidimacyname=split(Application("yx8_mhjh_illegidimacyname"),";")
regtime=now()
regip=Request.ServerVariables("REMOTE_ADDR")
password=md5(password)
on error resume next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open ("select 姓名 from 用户 where 姓名='"&username&"'"),conn
if not(rs.BOF or rs.EOF ) then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('对不起，此名字己被人注册！！！！');parent.history.go(-1);}</script>"
Response.End
end if 
conn.BeginTrans
conn.Execute "insert into 用户(帐号,密码,介绍人,头像,注册IP,注册时间,最后登录ip,最后登录时间,最后领钱,最后领钱日期,状态,形态,姓名,性别,门派,身份,配偶,存款,精力,等级,银两,积分,体力,内力,攻击,防御,资质,道德,风铃,特技,洞房花烛,结算日期,会员,会员时间,保护,protect,任务,任务时间,国家,职位,小妾,操作时间) values('"&account&"','"&password&"','"&jsr&"','"&txxx&"','"&regip&"','"&regtime&"','"&regip&"','"&regtime&"','"&regtime&"','"&regtime&"','正常','人','"&username&"','"&sex&"','无','无','无','0',200000,1,100000,0,1000000,100000,100000,100000,1000,1000,0,';','未','"&regtime&"',false,'"&regtime&"',true,'"&regtime&"','无','"&regtime&"','"&country&"','无','无','"&regtime&"')"
if conn.Errors.Count=0 then
conn.CommitTrans
else
conn.RollbackTrans
if conn.Errors(0).Number =-2147467259 then Response.Redirect "error.asp?id=009"
Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
end if
conn.Close
set conn=nothing
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【新人注册】</font><font color=blue>快乐江湖大家庭又增加新人口，<font color=red>"&username&"</font>成功注册，他所属的国家是"&country&"，大家要多照顾啊!谁先对"&username&"使用常用命令里的<font color=red>帮扶</font>功能,好处多多呢!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>

<HTML><HEAD><TITLE>战将注册</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="reg/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=40>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596>
  <TBODY> 
  <TR>
    <TD background=reg/history_table_bg.gif height=158 vAlign=top width="596"> 
      <div align="center">
        <center> 
      <TABLE border=10 cellPadding=0 cellSpacing=0 class=mountain 
      width=596 height="212" bordercolor="#840000" style="border-collapse: collapse; background-image:url('reg/page_bottom_mountain.gif')">
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="19" bordercolor="#000000">
          <table align="center" border="0" cellPadding="0" cellSpacing="0" width="596">
            <tr>
              <td background="reg/history_top_bg.gif" vAlign="top" width="596">
              <table align="center" border="0" cellPadding="2" cellSpacing="0" class="p9" width="97%" height="24">
                <tr>
                  <td height="20" width="31%">　</td>
                  <td height="20" vAlign="top" width="37%">
                  <div align="center">
                    <font color="#FF0000" size="3">魔幻战将注册</font></div>
                  </td>
                  <td height="20" vAlign="top" width="32%">　</td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </TD></TR> 
        <body leftmargin="0"  marginwidth="0" marginheight="0" bgcolor="#804040" topmargin="40">
<form action="regstep3.asp" method="post" id=form1 name=form1>

        <TR>
          <TD vAlign=top bgcolor="#000000" height="45" bordercolor="#000000">
          &nbsp;&nbsp;&nbsp;&nbsp;<align="center"><img border="0" src="reg/top1.gif">
            <div align="center">
              <center>
            <TABLE background=reg/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="537" height="53" style="border-collapse: collapse">
              <TBODY> 
              <TR>
          <div align="center"><font color="#ff0000"><b>注　册　成　功</b></font></div>
              </td>
            </tr>
            <tr> 
              <td width="50%" height="46">你的帐号：</td>
              <td width="50%" height="46"><%=trim(Request.Form("account"))%></td>
            </tr>
            <tr> 
              <td width="50%" height="46">你的姓名：</td>
              <td width="50%" height="46"><%=trim(Request.Form("username"))%></td>
            </tr>
            <tr> 
              <td width="50%" height="46">你的密码：</td>
              <td width="50%" height="46"><%=trim(Request.Form("password"))%></td>
            </tr>
            <tr> 
              <td width="50%" height="46">你的头像:</td>
              <td width="50%" height="46"><img src="<%=trim(Request.Form("txxx"))%>"></td>
            </tr>
            <tr> 
              <td colspan="2" height="36"> 
                <p align="center">&nbsp; 
                <p align="center">恭喜你成为快乐江湖超级战将！请牢记你的帐号、姓名和密码，以便登陆 
                <p align="center"> 
                  <input type='button' value='关闭' onClick='javascript:top.window.close();' id=button1 name=button1>
                           </td>
</td>       
            </tr>       
            <tr>        
</font></a></p>
                  </TD></p></TR></TBODY></TABLE>
              </center>
          </div>
            </TD></TR>

        <TR>
          <TD vAlign=top bgcolor="#000000" height="44" bordercolor="#000000">
          <p align="center"><img border="0" src="reg/top3.gif"></TD></TR></TBODY></TABLE></center>
      </div>
    </TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="reg/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>