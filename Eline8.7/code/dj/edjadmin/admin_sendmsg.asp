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
		errmsg="<li>����Ϣ����û����д���뷵����д������</li>"
    	call error()
		Response.End 
	elseif content="" then
		conn.close
		set conn=nothing
		errmsg="<li>����Ϣ����û����д���뷵����д������</li>"
    	call error()
		Response.End 
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	if request("touser")="tovip" then
	sql="select username from [user] where isvip=true and DateDiff('s',now(),vipenddate)>0"
	else
	sql="select username from [user] where lockuser<>true"
	end if
	rs.Open sql,conn,1,1
	if not rs.EOF then
		sendtime=now()
		sender="���ž���"
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
		msg="���ͳɹ�!"
	end if
	rs.Close 
	set rs=nothing
elseif action="del" then
	UserName=trim(Request.Form ("UserName"))
	DelR=trim(Request.Form ("delR"))
	DelW=trim(Request.Form ("delW"))
	if UserName="���ž���" then
		conn.execute("delete * from message")
		msg="�ɹ�ɾ�����ж���Ϣ"
	elseif DelR="on" or DelW="on" then
		if DelR="on" and DelW<>"on" then
			sql="delete * from message where incept='"&UserName&"'"
			thismsg="�յ�"
		elseif DelR<>"on" and DelW="on" then
			sql="delete * from message where sender='"&UserName&"'"
			thismsg="����"
		elseif DelR="on" and DelW="on" then
			sql="delete * from message where sender='"&UserName&"' or incept='"&UserName&"'"
			thismsg="����"
		end if
		conn.execute(sql)
		
		msg="�ɹ�ɾ���û�"&UserName&thismsg&"�Ķ���Ϣ"
	end if
elseif action="deldate" then
	thisdate=date()
	selectdate=cint(trim(Request.Form ("selectdate")))
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>"& selectdate &" and flag<>0")
	msg="�ɹ�ɾ��"&selectdate&"��ǰ�Ķ���Ϣ"
elseif action="delweek" then
	thisdate=date()
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>7 and flag<>0")
	msg="�ɹ�ɾ��һ��ǰ�Ķ���Ϣ"
elseif action="delmonth" then
	thisdate=date()
	conn.execute("delete * from message where datediff('d',datevalue(sendtime),#"& thisdate &"#)>30 and flag<>0")
	msg="�ɹ�ɾ��һ����ǰ�Ķ���Ϣ"
end if	
%>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="����.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign=top width=150>
    <!--#include file="admin_left.asp"-->
    ��</td>
    <td valign=top width=10>��</td>
    <td valign=top width="600">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#8DA2B7" width="100%" id="AutoNumber3">
      <tr>
        <td width="100%" bgcolor="#23213F">
      <p align="center"><b>���ž���</b></p>
      <div align="center">
        <center>
     <table border="1" width="90%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#666666">
       <tr>
         <td width="100%" height=22 align=center background="../images/tb_bg.gif"><b>
         �� ��</b></td>
       </tr>
       <tr>
         <td width=100% height=20>
           <li>���ž��齫�������ע���û�ͬʱ��������Ϣ��</li>
           <li>���ž�����ɾ��ĳ�������������˵Ķ���Ϣ������ɾ�����ڶ���Ϣ��</li>
           <li>����ɾ�������˵Ķ���Ϣ�������û��������롰���ž��顱�������ã�</li>
           <li>���ڶ���Ϣ������δ���Ķ���Ϣ��</li>
          
         </td>
       </tr>
       <tr>
         <td width="100%" align="center" colspan=2 background="../images/tb_bg.gif"><%=msg%>��</td>
       </tr>
     </table>
        </center>
      </div>
      <p><br>
      </p>
      <div align="center">
        <center>
     <table border="1" width="90%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#666666">
       <form method="POST" action="admin_sendmsg.asp?action=save" id=form1 name=form1>
         <tr>
           <td width="100%" height="20" colspan=2 align=center background="../images/tb_bg.gif"><b>����ͳһ����Ϣ</b></td>
         </tr>
         <tr>
           <td width="30%" align="right">* ����Ϣ���⣺</td>
           <td width="70%" height="30">
           <input type="text" name="title" size="35"><select size="1" name="touser">
             <option>���Ͷ���</option>
             <option value="tovip">VIP��Ա</option>
             <option value="toalluser">�����û�</option>
             </select></td>
         </tr>
         <tr>
           <td width="30%" align="right">* ����Ϣ���ݣ�</td>
           <td width="70%" height="30"><TEXTAREA name=content rows=6 cols=47></TEXTAREA></td>
         </tr>
         <tr>
           <td colspan=2 align=center background="../images/tb_bg.gif">
             <input type="submit" value=" ȷ�� " name="cmdok">&nbsp; 
             <input type="reset" value=" �� �� "  name="cmdcancel">
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
     <table border="1" width="50%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#666666">
       <form method="POST" action="admin_sendmsg.asp?action=del" id=form2 name=form2>
         <tr>
           <td width="100%" height="20" align=center background="../images/tb_bg.gif"><b>ɾ���û�����Ϣ</b></td>
         </tr>
         <tr>
           <td width="30%" align=center>* �û�����<input type="text" name="UserName" size="20"></td>
         </tr>
         <tr>
           <td width="30%" height="30" align=center>
           <input type="checkbox" name="delR" checked style='background-color:#c0c0c0; border: 0' value="ON">�����յ��Ķ���Ϣ</td>
         </tr>
         <tr>
           <td width="30%" height="30" align=center>
           <input type="checkbox" name="delW" checked style='background-color:#c0c0c0; border: 0' value="ON">���������Ķ���Ϣ</td>
         </tr>
         <tr>
           <td align=center background="../images/tb_bg.gif">
             <input type="submit" value=" ȷ�� " name="cmdok">&nbsp; 
             <input type="reset" value=" �� �� "  name="cmdcancel">
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
     <table border="1" width="50%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#666666" height="5">
         <tr>
           <td width="100%" height="19" align=center background="../images/tb_bg.gif"><b>ɾ�����ڶ���Ϣ</b></td>
         </tr>
       <form method="POST" action="admin_sendmsg.asp?action=deldate" id=form1 name=form1>
         <tr>
           <td width="30%" align=center height="19">���������<input type="text" name="selectdate" size="20"></td>
         </tr>
         <tr>
           <td align=center height="21">
             <input type="submit" value=" ȷ�� " name="cmdok">&nbsp; 
             <input type="reset" value=" �� �� "  name="cmdcancel">
           </td>
         </tr>
       </form>
         <tr>
           <td width="30%" height="1" align=center background="../images/tb_bg.gif">
           <a href="admin_sendmsg.asp?action=delweek">ɾ��һ��ǰ�Ķ���Ϣ</a></td>
         </tr>
         <tr>
           <td width="30%" height="1" align=center background="../images/tb_bg.gif">
           <a href="admin_sendmsg.asp?action=delmonth">ɾ��һ����ǰ�Ķ���Ϣ</a></td>
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
  </center>
</div>