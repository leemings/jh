<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>短信息管理</title>
</head>

<body>
<!--#include file="admin_top.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="checkadmin.inc"-->

<%
dim action,msg
action=Request.QueryString ("action")
if action="save" then
	title=trim(Request.Form ("title"))
	content=trim(Request.Form ("content"))
	if title="" then
		conn.close 
		set conn=nothing
		errmsg="<li>短信息标题没有填写！请返回填写完整。</li>"
    	call error()
		Response.End 
	elseif content="" then
		conn.close
		set conn=nothing
		errmsg="<li>短信息内容没有填写！请返回填写完整。</li>"
    	call error()
		Response.End 
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	sql="select username from [user] where username<>'短信精灵'"
	rs.Open sql,conn,1,1
	if not rs.EOF then
		sendtime=now()
		sender="短信精灵"
		do while not rs.EOF 
			sql="insert into message(incept,sender,title,content,sendtime) values('"
			sql=sql&(rs("username"))&"','"
			sql=sql&(sender)&"','"
			sql=sql&(title)&"','"
			sql=sql&(content)&"','"
			sql=sql&(sendtime)&"')"
			conn.execute(sql)
		rs.MoveNext 
		loop
	end if
	if err then
		conn.close
		set conn=nothing
		Response.Write err.description
		Response.End 
	else
		msg="发送成功!"
	end if
	rs.Close 
	set rs=nothing
elseif action="del" then
	UserName=trim(Request.Form ("UserName"))
	DelR=trim(Request.Form ("delR"))
	DelW=trim(Request.Form ("delW"))
	if UserName="短信精灵" then
		conn.execute("delete * from message")
		msg="成功删除所有短信息"
	elseif DelR="on" or DelW="on" then
		if DelR="on" and DelW<>"on" then
			sql="delete * from message where incept='"&UserName&"'"
			thismsg="收到"
		elseif DelR<>"on" and DelW="on" then
			sql="delete * from message where sender='"&UserName&"'"
			thismsg="发出"
		elseif DelR="on" and DelW="on" then
			sql="delete * from message where sender='"&UserName&"' or incept='"&UserName&"'"
			thismsg="所有"
		end if
		conn.execute(sql)
		
		msg="成功删除用户"&UserName&thismsg&"的短信息"
	end if
elseif action="deldate" then
	thisdate=date()
	selectdate=cint(trim(Request.Form ("selectdate")))
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>"& selectdate &" and flag<>0")
	msg="成功删除"&selectdate&"天前的短信息"
elseif action="delweek" then
	thisdate=date()
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>7 and flag<>0")
	msg="成功删除一周前的短信息"
elseif action="delmonth" then
	thisdate=date()
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>30 and flag<>0")
	msg="成功删除一个月前的短信息"
end if	
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="补间.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse" height="155">
  <tr>
    <td valign=top width=150 height="155">
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10 height="155">　</td>
    <td valign=top width="600" height="155">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" height="100%" bordercolor="#000000">
      <tr>
        <td width="100%" bgcolor="#FF9933" valign="top" height="84">
    <table border="0" cellpadding="3" cellspacing="3" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%">
      <p align="center"><b>短信精灵</b></p>
      <div align="center">
        <center>
     <table border="1" width="90%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#CC6600">
       <tr>
         <td width="100%" height=22 align=center bgcolor="#CC6600"><b>
         帮 助</b></td>
       </tr>
       <tr>
         <td width=100% height=20>
           <li>短信精灵将会对所有注册用户同时发出短信息。</li>
           <li>短信精灵能删除某个人甚至所有人的短信息，还能删除过期短信息。</li>
           <li>如想删除所有人的短信息，请在用户栏内输入“短信精灵”。请慎用！</li>
           <li>过期短信息不包括未读的短信息。</li>
          
         </td>
       </tr>
       <tr>
         <td width="100%" align="center" colspan=2 bgcolor="#CC6600"><%=msg%>　</td>
       </tr>
     </table>
        </center>
      </div>
      <p><br>
      </p>
      <div align="center">
        <center>
     <table border="1" width="90%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#CC6600">
       <form method="POST" action="sendmsg.asp?action=save" id=form1 name=form1>
         <tr>
           <td width="100%" height="20" colspan=2 align=center bgcolor="#CC6600"><b>发送统一短信息</b></td>
         </tr>
         <tr>
           <td width="30%" align="right">* 短信息标题：</td>
           <td width="70%" height="30">
           <input type="text" name="title" size="35"></td>
         </tr>
         <tr>
           <td width="30%" align="right">* 短信息内容：</td>
           <td width="70%" height="30"><TEXTAREA name=content rows=6 cols=47></TEXTAREA></td>
         </tr>
         <tr>
           <td colspan=2 align=center bgcolor="#CC6600">
             <input type="submit" value=" 确定 " name="cmdok">&nbsp; 
             <input type="reset" value=" 清 除 "  name="cmdcancel">
           </td>
         </tr>
       </form>
     </table>
        </center>
      </div>
      <p><br>
      </p>
      <div align="center">
        <center>
     <table border="1" width="50%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#CC6600">
       <form method="POST" action="sendmsg.asp?action=del" id=form2 name=form2>
         <tr>
           <td width="100%" height="20" align=center bgcolor="#CC6600"><b>删除用户短信息</b></td>
         </tr>
         <tr>
           <td width="30%" align=center>* 用户名：<input type="text" name="UserName" size="20"></td>
         </tr>
         <tr>
           <td width="30%" height="30" align=center>
           <input type="checkbox" name="delR" checked style='background-color:#c0c0c0; border: 0' value="ON">包含收到的短信息</td>
         </tr>
         <tr>
           <td width="30%" height="30" align=center>
           <input type="checkbox" name="delW" checked style='background-color:#c0c0c0; border: 0' value="ON">包含发出的短信息</td>
         </tr>
         <tr>
           <td align=center bgcolor="#CC6600">
             <input type="submit" value=" 确定 " name="cmdok">&nbsp; 
             <input type="reset" value=" 清 除 "  name="cmdcancel">
           </td>
         </tr>
       </form>
     </table>
        </center>
      </div>
      <p><br>
      </p>
      <div align="center">
        <center>
     <table border="1" width="50%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#CC6600" height="5">
         <tr>
           <td width="100%" height="19" align=center bgcolor="#CC6600"><b>删除过期短信息</b></td>
         </tr>
       <form method="POST" action="sendmsg.asp?action=deldate" id=form1 name=form1>
         <tr>
           <td width="30%" align=center height="19">间隔天数：<input type="text" name="selectdate" size="20"></td>
         </tr>
         <tr>
           <td align=center height="21">
             <input type="submit" value=" 确定 " name="cmdok">&nbsp; 
             <input type="reset" value=" 清 除 "  name="cmdcancel">
           </td>
         </tr>
       </form>
         <tr>
           <td width="30%" height="1" align=center bgcolor="#CC6600">
           <a href="sendmsg.asp?action=delweek">删除一周前的短信息</a></td>
         </tr>
         <tr>
           <td width="30%" height="1" align=center bgcolor="#CC6600">
           <a href="sendmsg.asp?action=delmonth">删除一个月前的短信息</a></td>
         </tr>
     </table>
        </center>
      </div>
</td>
      </tr>
    </table>

</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>
</body>

</html>