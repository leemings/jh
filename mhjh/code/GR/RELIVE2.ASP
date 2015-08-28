<!--#include file="pass.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,"?")<>0 or instr(elevalue,";") or instr(elevalue,"file:")<>0 or instr(elevalue,"script:")<>0 or instr(elevalue,"#")<>0 then Response.Redirect "../error.asp?id=056"
next
account=trim(Request.Form("account"))
username=trim(Request.Form("username"))
password=trim(Request.Form("password"))
if account="" or username="" or password="" then Response.Redirect "../error.asp?id=002"
if len(account)<6 or len(username)<2 or len(password)<6 then Response.Redirect "../error.asp?id=004"
if len(account)>15 or len(username)>7 or len(password)>15 then Response.Redirect "../error.asp?id=504"
if instr(account,"file:")<>0 or instr(account,"script:")<>0 or instr(account,"'")<>0 or instr(account,chr(34))<>0 or instr(account,"\")<>0 or InStr(account,",")<>0 or InStr(account,"<")<>0 or InStr(account,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(username,"file:")<>0 or instr(username,"script:")<>0 or instr(username,"'")<>0 or instr(username,chr(34))<>0 or instr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then Response.Redirect "../error.asp?id=524"
if instr(password,"file:")<>0 or instr(password,"script:")<>0 or instr(password,"'")<>0 or instr(password,chr(34))<>0 or instr(password,"\")<>0 or InStr(password,",")<>0 or InStr(password,"<")<>0 or InStr(password,">")<>0 then Response.Redirect "../error.asp?id=524"
for i=1 to len(account)
accountchr=asc(mid(account,i,1))
if accountchr<48 or (accountchr>57 and accountchr<65) or (accountchr>90 and accountchr<97) then Response.Redirect "../error.asp?id=006"
next
for i=1 to len(username)
usernamechr=mid(username,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=003"
next
for i=1 to len(password)
passwordchr=asc(mid(password,i,1))
if passwordchr<48 or (passwordchr>57 and passwordchr<65) or (passwordchr>90 and passwordchr<97) or passwordchr>122 then Response.Redirect "../error.asp?id=005"
next
password=md5(password)
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 会员,等级,状态,最后登录时间 from 用户 where 帐号='"&account&"' and 姓名='"&username&"' and 密码='"&password&"'",conn
if rst.EOF or rst.BOF  then Response.Redirect "../error.asp?id=019"
if rst("状态")<>"死亡" then Response.Redirect  "../error.asp?id=020"
on error resume next
if rst("会员")=True then
conn.Execute "update 用户 set 状态='正常',最后登录时间='" & now() & "' where 姓名='"&username&"' and 密码='"&password&"' and 帐号='"&account&"' and 状态='死亡'"
else
conn.Execute "update 用户 set 状态='正常',形态='鬼魂',门派='十八地狱',最后登录时间='" & now() & "' where 姓名='"&username&"' and 密码='"&password&"' and 帐号='"&account&"' and 状态='死亡'"
end if
if conn.Errors.Count=0 then
conn.CommitTrans
else
conn.RollbackTrans
Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
end if
set rst=nothing
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
		newtalkarr(599)="<font color=red>【战将复活】</font><b><font color=red>请大家注意："&username&"复活失败，变成了冤魂野鬼，在十八地狱里正苦苦挣扎，如果这个鬼魂熬过奈何桥，他将舍命报酬，杀人凶手要小心了。</font></b>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>

<HTML><HEAD><TITLE>战将复活</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="reg/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=30>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596>
  <TBODY> 
  <TR>
    <TD background=reg/history_table_bg.gif height=158 vAlign=top width="596"> 
      <div align="center">
        <center> 
      <TABLE border=10 cellPadding=0 cellSpacing=0 class=mountain 
      width=596 height="153" bordercolor="#840000" style="border-collapse: collapse; background-image:url('reg/page_bottom_mountain.gif')">
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="19">
          <table align="center" border="0" cellPadding="0" cellSpacing="0" width="596">
            <tr>
              <td background="reg/history_top_bg.gif" vAlign="top" width="596">
              <table align="center" border="0" cellPadding="2" cellSpacing="0" class="p9" width="97%" height="24">
                <tr>
                  <td height="20" width="31%">　</td>
                  <td height="20" vAlign="top" width="37%">
                  <div align="center">
                    <font color="#FF0000" size="3">战将复活</font></div>
                  </td>
                  <td height="20" vAlign="top" width="32%">　</td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="30" bordercolor="#000000">
          <p align="center"><img border="0" src="reg/top1.gif"></TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="45" bordercolor="#000000">
            <div align="center">
              <center>
            <TABLE background=reg/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="537" height="232" style="border-collapse: collapse">
              <TBODY> 
                <tr align="center" valign="middle"> 
      <td height="97" width="537">　<p align="center"><font color="#FFFF00">会员复活不会变成鬼魂,非会员要变成鬼魂，<br>
      每次登陆江湖要减少体力的5%，其他状态不会丢失。尽快凤凰蘖盘那进化成人吧！</font><br><font color="#FFFFFF">你已经获得重生，快去
      地狱</font><font color="#FF0000">凤凰蘖盘</font><font color="#FFFFFF">修炼成人<br>好好练功，找你的仇人报仇吧</font></p>
<p align="center"><a href="relive2.asp#" onclick="javascript:window.close();">关闭</a></p>
</td>
    </tr>
    <form action="relive2.asp" method="post">
              </TBODY></TABLE>
              </center>
          </div>
            </TD></TR> 
        <TR>
          <TD vAlign=top bgcolor="#000000" height="44">
          <p align="center"><img border="0" src="reg/top3.gif"></TD></TR></TBODY></TABLE></center>
      </div>
    </TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="reg/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>