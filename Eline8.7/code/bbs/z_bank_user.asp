<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<%
dim MaxAnnouncePerPage,rs1
stats="电子银行 储户管理"
call nav()
call head_var(0,0,"电子银行","Z_bank.asp")
if not master then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请登录后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
   <tr> 
       <th  height=25>欢迎<b><%=membername%></b>进入银行管理页面</th>
   </tr> 
   <tr> 
       <td width="100%" class=tablebody2> 

<% 
	call bankhead()
	
	MaxAnnouncePerPage=Forum_Setting(11) 
   	dim totalPut    
   	dim CurrentPage 
   	dim TotalPages 
   	dim idlist 
   	dim title 
   	dim sql1,rs11 
	title=request("txtitle") 
   	if not isempty(request("page")) then 
      		currentPage=cint(request("page")) 
  	else 
      		currentPage=1 
   	end if 
%>
<form name="searchuser" method="POST" action="Z_bank_user.asp">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">
<tr><td class=tablebody2> 
 	<font color=red>点击用户名进行相应操作</font>，  查找用户:  <input type="text" name="txtitle" size="13">	<input type="submit" value="查询" name="title">                                            
</td></tr>
</table>                                      
</form>                                         
<%                                          
	if request("action")="del" then                                          
		call del()                                          
	end if                                          
                                         
	if title<>"" then                                          
		sql1="select *  from [bank] where username like '%"&checkStr(trim(title))&"%' order by username desc"                                          
	else                                          
		sql1="select * from [bank] where username <> '银行' order by savemoney desc"                                          
	end if                                          
	Set rs1= Server.CreateObject("ADODB.Recordset")                                          
	rs1.open sql1,conn1,1,1                                          
                                          
  	if rs1.eof and rs1.bof then                                          
       		response.write "<p align='center'> 还 没 有 任 何 用 户 </p>"                                          
   	else                                          
      		totalPut=rs1.recordcount                                           
      		if currentpage<1 then                                           
          		currentpage=1                                           
      		end if                                           
			MaxAnnouncePerpage=Clng(MaxAnnouncePerpage)                                          
      		if (currentpage-1)*MaxAnnouncePerPage>totalput then                                           
				if (totalPut mod MaxAnnouncePerPage)=0 then                                           
						currentpage= totalPut \ MaxAnnouncePerPage                                           
				else                                           
						currentpage= totalPut \ MaxAnnouncePerPage + 1                                           
				end if                                           
      		end if                                           
       		if currentPage=1 then                                           
            		showContent                                           
            		showpage totalput,MaxAnnouncePerPage,"Z_bank_user.asp"                                           
       		else                                           
          		if (currentPage-1)*MaxAnnouncePerPage<totalPut then                                           
            			rs1.move  (currentPage-1)*MaxAnnouncePerPage                                          
            			dim bookmark                                           
            			bookmark=rs1.bookmark                                            
            			showContent                                           
             			showpage totalput,MaxAnnouncePerPage,"Z_bank_user.asp"                                           
        		else                                           
	        		currentPage=1                                           
           			showContent                                           
           			showpage totalput,MaxAnnouncePerPage,"Z_bank_user.asp"                                           
	      		end if                                           
	   		end if                                          
   	end if 
	rs1.close                                         
	                                                  
   	set rs1=nothing                                            
   	conn1.close                                          
   	set conn1=nothing
%>
	</td></tr>
	<tr><td align="center" class="tablebody2"><a href=Z_bank.asp>返回银行首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>返回上一页</a></td></tr>
	</table>
<%	
 end if
  	                                          
 call activeonline()	
 call footer() 
'=====================================================================     
                                     
sub showContent                                          
   	dim i                                          
	i=0                                          
%>                                          
<table cellspacing="1" cellpadding="3" class=tableborder1 align="center" style="width:97%">                                          
        <tr>                                          
                    <th height="20">用户名</th>                                                    
                    <th>存款</th>                                          
					<th>存款时间</th>                                        
                    <th>利息</th>                                          
				    <th>未还贷款</th>                                          
					<th>贷款时间</th>                                          
					<th width="65">冻结帐户</th>
					<th>操作</th>                                         
        </tr>                                          
<%
	dim saveday
	do while not rs1.eof
		saveday=datediff("d",rs1("date"),date())
%>                                          
        <tr>                                          
                    <td class=tablebody2 height="23" align=center><a href="Z_bank_acuser.asp?name=<%=(rs1("username"))%>"><%=(rs1("username"))%></a></td>                                          
                    <td class=tablebody1 align=center><%=rs1("savemoney")%></td>                                          
					<td class=tablebody2 align=center><%=rs1("date")%></td>                                          
                    <td class=tablebody1 align=center><%if saveday<Chen_StartLixiDay then %>0<%else%><font color="#0066FF"><%=clng((formatnumber(rs1("savemoney"))*(formatnumber(culi)/100))*saveday)%></font><%end if%></td>                                          
                    <td class=tablebody2 align=center><%=rs1("daikuang")%></td>                                          
					<td class=tablebody1 align=center><%=rs1("dkdate")%> </td>  
					<td class=tablebody2 align=center><%if rs1("lockac") then%><font color="#FF0000">是</font><%else%>否<%end if%></td>                                         
                    <td class=tablebody1 align=center><a href="Z_bank_user.asp?action=del&name=<%=rs1("username")%>" onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}">删除</a></td>                                          
        </tr>                                          
                                          
<%                                           
	i=i+1                                          
	if i>=MaxAnnouncePerPage then exit do                                          
	rs1.movenext                                          
	loop                                          
%>                                          
      </table>                                          
<%                                          
end sub                                           
                                          
function showpage(totalnumber,MaxAnnouncePerPage,filename)                                          
    dim n                                    
  if totalnumber mod MaxAnnouncePerPage=0 then                                          
     n= totalnumber \ MaxAnnouncePerPage                                          
  else                                          
     n= totalnumber \ MaxAnnouncePerPage+1                                          
  end if                                          
  response.write "<p align='center'>&nbsp;"                                          
  if CurrentPage<2 then                                          
    response.write "首页 上一页&nbsp;"                                          
  else                                          
    response.write "<a href="&filename&"?page=1&txtitle="&request("txtitle")&">首页</a>&nbsp;"                                          
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&txtitle="&request("txtitle")&">上一页</a>&nbsp;"                                          
  end if                                          
  if n-currentpage<1 then                                          
    response.write "下一页 尾页"                                          
  else                                          
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&txtitle="&request("txtitle")&">"                                          
    response.write "下一页</a> <a href="&filename&"?page="&n&"&txtitle="&request("txtitle")&">尾页</a>"                                          
  end if                                          
   response.write "&nbsp;页次：</font><strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>页 "                                          
   response.write "&nbsp;共<b>"&totalnumber&"</b>个用户 <b>"&MaxAnnouncePerPage&"</b>个用户/页 "                                          
                                                 
end function                                          
                                          
sub del()   
	dim username                                       
    username=checkStr(trim(request.querystring("name")))
    sql1="delete from bank where username='"&username&"'"                                          
	conn1.Execute(sql1)                                          
	call savemsg(username)
	
	sucmsg=""
	if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
		content="注销用户[<font color=navy>"&username&"</font>]的银行账号"
		call logs("管理","储户管理",membername)
		sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
	end if	                         
%>                         
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>银行操作成功</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%>电子银行
	<font color=navy><%=sucmsg%><li>成功删除用户[<font color=red><%=username%></font>]的银行账号<br><li>并发送了"银行紧急通知"给该用户</font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>返回上一页</a></td></tr>
</table>
<%
end sub
                                  
sub savemsg(username)  
	dim message                                        
	title="银行紧急通知"                                          
	message=username&"你好！由于你有扰乱银行的行为，你在社区银行里的所有数据已由银行工作员清空，如有问题请到社区办公室询问"                                          
	sql1="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&username&"','"&"电子银行办公室"&"','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql1)                                          
end sub                                          
%> 