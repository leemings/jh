<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<%
dim MaxAnnouncePerPage,rs1
stats="�������� ��������"
call nav()
call head_var(0,0,"��������","Z_bank.asp")
if not master then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã����¼����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
   <tr> 
       <th  height=25>��ӭ<b><%=membername%></b>�������й���ҳ��</th>
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
 	<font color=red>����û���������Ӧ����</font>��  �����û�:  <input type="text" name="txtitle" size="13">	<input type="submit" value="��ѯ" name="title">                                            
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
		sql1="select * from [bank] where username <> '����' order by savemoney desc"                                          
	end if                                          
	Set rs1= Server.CreateObject("ADODB.Recordset")                                          
	rs1.open sql1,conn1,1,1                                          
                                          
  	if rs1.eof and rs1.bof then                                          
       		response.write "<p align='center'> �� û �� �� �� �� �� </p>"                                          
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
	<tr><td align="center" class="tablebody2"><a href=Z_bank.asp>����������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>������һҳ</a></td></tr>
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
                    <th height="20">�û���</th>                                                    
                    <th>���</th>                                          
					<th>���ʱ��</th>                                        
                    <th>��Ϣ</th>                                          
				    <th>δ������</th>                                          
					<th>����ʱ��</th>                                          
					<th width="65">�����ʻ�</th>
					<th>����</th>                                         
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
					<td class=tablebody2 align=center><%if rs1("lockac") then%><font color="#FF0000">��</font><%else%>��<%end if%></td>                                         
                    <td class=tablebody1 align=center><a href="Z_bank_user.asp?action=del&name=<%=rs1("username")%>" onclick="javascript:{if(confirm('��ȷ��ִ��ɾ��������?')){return true;}return false;}">ɾ��</a></td>                                          
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
    response.write "��ҳ ��һҳ&nbsp;"                                          
  else                                          
    response.write "<a href="&filename&"?page=1&txtitle="&request("txtitle")&">��ҳ</a>&nbsp;"                                          
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&txtitle="&request("txtitle")&">��һҳ</a>&nbsp;"                                          
  end if                                          
  if n-currentpage<1 then                                          
    response.write "��һҳ βҳ"                                          
  else                                          
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&txtitle="&request("txtitle")&">"                                          
    response.write "��һҳ</a> <a href="&filename&"?page="&n&"&txtitle="&request("txtitle")&">βҳ</a>"                                          
  end if                                          
   response.write "&nbsp;ҳ�Σ�</font><strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "                                          
   response.write "&nbsp;��<b>"&totalnumber&"</b>���û� <b>"&MaxAnnouncePerPage&"</b>���û�/ҳ "                                          
                                                 
end function                                          
                                          
sub del()   
	dim username                                       
    username=checkStr(trim(request.querystring("name")))
    sql1="delete from bank where username='"&username&"'"                                          
	conn1.Execute(sql1)                                          
	call savemsg(username)
	
	sucmsg=""
	if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
		content="ע���û�[<font color=navy>"&username&"</font>]�������˺�"
		call logs("����","��������",membername)
		sucmsg=sucmsg+"<br>"+"<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ�"
	end if	                         
%>                         
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���в����ɹ�</th></tr>
	<tr><td width="100%" class=tablebody1><b>�����ɹ���</b><br><br><li>��ӭ����<%=Forum_info(0)%>��������
	<font color=navy><%=sucmsg%><li>�ɹ�ɾ���û�[<font color=red><%=username%></font>]�������˺�<br><li>��������"���н���֪ͨ"�����û�</font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>������һҳ</a></td></tr>
</table>
<%
end sub
                                  
sub savemsg(username)  
	dim message                                        
	title="���н���֪ͨ"                                          
	message=username&"��ã����������������е���Ϊ��������������������������������й���Ա��գ����������뵽�����칫��ѯ��"                                          
	sql1="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&username&"','"&"�������а칫��"&"','"&title&"','"&message&"',Now(),0,1)"                                          
	conn.execute(sql1)                                          
end sub                                          
%> 