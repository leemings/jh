<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%	

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
function wpadd(wpadd_name,wpadd_user,wpadd_sl)	'物品名，账号，数量，类别
	wpadd=0
	Set conn_wp = Server.CreateObject("ADODB.CONNECTION")
	Set rs_wp = Server.CreateObject("ADODB.RecordSet")
	conn_wp.Open Application("aqjh_usermdb")
	rs_wp.open "select * from wpname where wp_user='" & wpadd_user & "' and wp_name='"& wpadd_name &"' and wp_sl>0" ,conn_wp,1,3
	If rs_wp.Bof and rs_wp.Eof then
		rs_wp.close
		'查找数据库中已删除记录使用删除亡不保存新物品
		rs_wp.open "select * from wpname where wp_sl=0" ,conn_wp,1,3
		If rs_wp.Bof and rs_wp.Eof then
			rs_wp.addnew
		end if
		wpadd=wpadd_sl				'返回数量
		rs_wp("wp_sl")=clng(wpadd_sl)
	else
		wpadd=rs_wp("wp_sl")+clng(wpadd_sl)	'返回数量
		rs_wp("wp_sl")=rs_wp("wp_sl")+clng(wpadd_sl)
	end if
	rs_wp("wp_user")=wpadd_user
	rs_wp("wp_name")=wpadd_name
	rs_wp.update
	rs_wp.close
	set rs_wp=nothing
	conn_wp.close
	set conn_wp=nothing
end function

err_num=0
err_info=""
info=""
Redir=""
dy=session("dy")
if IsArray(dy) then
	if isnull(dy(0)) then
		err_num=err_num+1
		err_info="没账号"
	else
		myid=dy(0)
	end if
else
	err_num=err_num+1
	err_info="你没开始钓鱼啊"
end if
if err_num>0 then
	Response.Write(err_info)
	Response.end
end if

Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open Application("aqjh_usermdb")
Set Rs = Server.CreateObject("ADODB.Recordset")
if dy(8)=""  then
	'--------------------------产生物品
	wp_sort="渔"
	Sql_fen_wp="SELECT * from b WHERE b = '" & wp_sort & "' and  p LIKE '%" & dy(2) & ";%'"
	Rs.Open Sql_fen_wp,Conn,1,1
	if Rs.Bof and Rs.Eof then	
		err_num=err_num+1
		info="没找到物品"
		session.Contents.Remove("dy")
		Rs.Close()
	else
		Randomize   '初始化随机数生成器。
		nn=Int(Rs.RecordCount * Rnd)
		Rs.MoveFirst
		Rs.move(nn)
		wp_name=(Rs.Fields.Item("a").Value)'名称
		'wp_j=(Rs.Fields.Item("j").Value)'价钱
		wp_sm=(Rs.Fields.Item("c").Value)'说明
		wp_i=(Rs.Fields.Item("i").Value)'图片
		wp_j=(Rs.Fields.Item("j").Value)'重量
		Rs.Close()
		if IsNull(wp_i) or wp_i="" or len(wp_i)=0 then wp_i="2.gif"
		dy(8)=wp_name'物品编号
		dy(9)=cint(wp_j*.1+wp_j*.9*Rnd)'物品数量
		session("dy")=dy
		Conn.Execute("update 用户 set jh_fish=jh_fish+1 where 姓名='"&dy(0)&"'") 
		info="得到物品"&wp_name&"重量"&dy(9)
		end if
else
	myid=dy(0)
	wp_name=dy(8)
	wp_num=dy(9)

	Sql_wp="SELECT * from b WHERE a = '" & wp_name & "'"
	Rs.Open Sql_wp,Conn,1,1
	if Rs.Bof and Rs.Eof then 
		info="非法物品"
		err_num=err_num+1
		session.Contents.Remove("dy")
	else
		wp_name=(Rs.Fields.Item("a").Value)'名称
		wp_j=(Rs.Fields.Item("h").Value)'价钱
		wp_jb=(Rs.Fields.Item("m").Value)'价钱
		wp_sm=(Rs.Fields.Item("c").Value)'说明
		wp_i=(Rs.Fields.Item("i").Value)'图片
		'wp_j=(Rs.Fields.Item("j").Value)'重量
		if IsNull(wp_i) or wp_i="" or len(wp_i)=0 then wp_i="2.gif"
	end if
	Rs.Close()
	info=myid&"处理物品"&wp_name&"数量"&wp_num&"   "&(wp_j*wp_num)&"银两钱"&(wp_jb*wp_num)&"金币"
'Response.Write(info)
	deal_way=Trim(Request.Form("deal_way"))
	if err_num=0 and (deal_way="1" or deal_way="2") then '处理物品
		'---------------------------------------卖东西
		if deal_way="1" then 
			'Response.Write(sql)
			Sql_3="SELECT * FROM 用户 WHERE 姓名 = '"& myid &"'"
			Rs.Open Sql_3,Conn,0,1
			If Rs.Bof and Rs.Eof Then '没找到ren
			Else '找到 物品存在修改物品数量
				sql="update 用户 set 金币=金币+"&(wp_jb*wp_num)&", 银两=银两+"&(wp_j*wp_num)&" where 姓名='"&myid&"'"
				Conn.Execute (sql)
				info="你把"&wp_name&"卖了,得到"&(wp_j*wp_num)&"银两钱"&(wp_jb*wp_num)&"金币"
				info1="钓到一条"&wp_name&",并把"&wp_name&"卖了,得到"&(wp_j*wp_num)&"银两钱"&(wp_jb*wp_num)&"金币"
				Redir="../index.asp?info="&info
				session.Contents.Remove("dy") '结束钓鱼
			End If 
			Rs.Close()
		end if
		'---------------------------------------把东西放到仓库中
		if deal_way="2" then
				a=wpadd(wp_name,myid,wp_num)
				session.Contents.Remove("dy") '结束钓鱼
				info="你把"&wp_num&"斤"&wp_name&"放到仓库中"
				info1="钓到一条"&wp_num&"斤"&wp_name&",并把"&wp_num&"斤"&wp_name&"放到仓库中"
				Redir="../index.asp?info="&info
		End If
	end if

says="<font color=#ff0000>【钓鱼信息】" & aqjh_name & ""& info1 &"</font>"			'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
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
end if



Set Rs = Nothing
Conn.Close()
Set Conn = Nothing
'if Redir<>"" then Response.Redirect(Redir)
if Redir<>"" then 
Response.Write("<script language=""JavaScript"">alert('提示："&info&"');window.close()</script>")
Response.end
end if
%>

<html>
<head>
<html>
<head>
<title>渔人码头</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META HTTP-EQUIV="Cache-Control:" CONTENT="no-cache, must-revalidate">
<link href="../the9.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
var isSubmited = 0;
function formSubmit()
{
  if (isSubmited)	{
 		alert("手指不要抖啊，回家看看用户事件，您已经做过了。");
		return;
  }
  else {
	 isSubmited = 1;
     
  }
  document.deal.ok.value=1;
  document.deal.submit();
}
</script>
</head>
<body bgcolor="#ffffff" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0"> 
<table width="760" border=0 cellpadding="0" cellspacing="0" bgcolor="#E8E8F0">
  <tr>
    <td width="167" valign="top"><br>
        <table width="167" border="0" cellspacing="0" cellpadding="0" height="80">
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/cloud02.gif" width="60" height="24"></td>
            <td><br>
                <br>
                <img src="../images/cloud03.gif" width="34" height="14"></td>
          </tr>
      </table></td>
    <td> <img src="../images/ship01.gif" width="129" height="46"><img src="../images/ship02.gif" width="137" height="95" border="0"></td>
    <td valign=top><img src="../images/ship03.gif" width="50" height="60"></td>
    <td width="100%">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="90">
        <tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/cloud02.gif" width="60" height="24"></td>
          <td valign="bottom"><img src="../images/cloud03.gif" width="33" height="14"><br>
              <br></td>
          <td><img src="../images/cloud01.gif" width="69" height="28"><br />
              <br></td>
        </tr>
    </table></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" bgcolor="#d3d7ea"> 
  <tr> 
    <td><br></td> 
  </tr> 
  <tr> 
    <td align="center"> <table width="760" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td width="20">　</td> 
          <td width="120"> <table border="0" cellspacing="0" cellpadding="0" align="center"> 
              <tr> 
                <td align="right"><img src="../images/a1.gif" width="18" height="312"></td> 
                <td><img src="../images/<% Response.Write(wp_i) %>" border=0 alt="<%= wp_name %>"  width="325" height="312"></td> 
                <td><img src="../images/a2.gif" width="53" height="312"></td> 
              </tr> 
            </table></td> 
          <td width="20">　</td> 
          <td width="300" valign="top" align=right> <br> 
            <table width="300" border="0" cellspacing="0" cellpadding="0" class="ct-def1" > 
              <tr> 
                <td class=ct-def1>您钓到的鱼是<font color=#cc3366><%= wp_name %></font>，重<font color=#cc3366><b><%= dy(9) %></b></font>斤。<br> 
                （点击鱼儿的图片可以看到该鱼类的介绍哦）</td> 
              </tr> 
              <tr> 
                <td class=ct-def1><font color=#cc3366><%= wp_sm %></font></td> 
              </tr> 
              <tr> 
                <td><br></td> 
              </tr> 
              <form action="" method="post" name=deal> 
                <tr> 
                  <td class=ct-def1>菜场收购价是：<%= wp_j %>/斤</td> 
                </tr> 
                <tr> 
                  <td><br></td> 
                </tr> 
                <tr> 
                  <td class=ct-def1>打算怎么处理它呢？</td> 
                </tr> 
                <tr> 
                  <td class=ct-def1> <input type="radio" name="deal_way" value="1"> 
                    卖给菜场<br> 
         <%    '       <input type="radio" name="deal_way" value="2"> %>
                    放进菜篮子         (功能开发中)<br> </td> 
                </tr> 
                <tr> 
                  <td><br></td> 
                </tr> 
                <tr> 
                  <td align="right"><a href="javascript:formSubmit()" class=ct-def1><img src="../images/ok.gif" border=0 width=36 height=36 align="absmiddle">确定</a> 
                    <input type="hidden" name="ok" value="0"></td> 
                </tr> 
              </form> 
              <tr> 
                <td colspan=2><br></td> 
              </tr> 
            </table></td> 
        </tr> 
      </table> 
  <tr> 
    <td><br></td> 
  </tr> 
  </td> 
  </tr> 
</table> 

</body>
</html>

</body>
</html>
