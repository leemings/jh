<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<script>
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script> 
<script language="JavaScript" type="text/JavaScript">
function show(n)
{
if (document.all["p_group"][n].style.display=="")
{

	document.all["p_group"][n].style.display="none"
	document.all["f_flag"][n].innerHTML="6";
}
else
{
	document.all["f_flag"][n].innerHTML="5";
	document.all["p_group"][n].style.display=""
}
}

function phoidall(form)  {
  for (var i=0;i<form.abc.length;i++)    {
    var e = form.abc[i];
    if (e.name != 'pho_id_all') e.checked = form.pho_id_all.checked; 
   }
  }
</script>


<% '	if not master or session("flag")="" then
	'Response.Write "<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"		
	'else
	'	Response.Write "you are master"
	'end if %>
<%


dim menu

menu=request.form("menu")
if menu="" then menu=0
stats="�������"

call nav()

call head_var(2,0,"","")

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н����������Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call start()
end if
if founderr then  call dvbbs_error()
call activeonline()
call footer() 
'=========================================================
sub start()

 select case menu
		case 0
			'call del()					'Ĭ��
		case 1
			call m_del()				'ɾ����Ƭ

		case 2
			call m_move()				'ɾ����Ƭ		case else 
														
	end select 

'-------------------------
%>
<table width="97%" align=center cellpadding=3 cellspacing=1 class=tableborder1>
	<tr>
		
    <th height=25>�� �� �� ��</td> </tr>
	<tr>
		<td class=tablebody2 height=1>	

<br>
            <form name="form1" method="get" action="z_photo3.asp">
        <table width="98%" height="25" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tr> 
            <td align=left valign=middle class=TableTitle2><font face="Webdings">2</font>&nbsp;<font color="BLUE"><a href="z_photo.asp">�����ҳ</a></font> 
              <font face="Webdings">2</font><a href="z_photo4.asp"><font color="blue">&nbsp;</font>����б�</a> 
              <font face="Webdings">2</font>&nbsp;<font color=blue><a href="z_photo_group.asp">��������</a></font> 
              <font face="Webdings">2</font>&nbsp;<a href="z_photo2.asp">��Ƭ����</a> 
              <%if master then %>&nbsp;<font face="Webdings">2</font>&nbsp;<a href="z_photo_admin.asp">�ܹ���</a><% end if %></td>
            <td align=right valign=middle class=TableTitle2>�û����� 
              <input name="pho_names" type="text" id="pho_names" size="16"> <input type="submit" name="Submit" value="����������"> 
            </td>
          </tr>
        </table>
        <br>
		<% 
		dim pho_mx,pho_mx_sh,pho_mx_all,ms
	set ms=server.createobject("adodb.recordset")
    	sql="select count(*) as pho_mx from z_photo_group where (pho_state = 'v' )"
    ms.open sql,z_photo_cn,1,3
	pho_mx=ms("pho_mx")
	ms.close
	    sql="select count(*) as pho_mx_all from z_photo_group where (pho_state = 'v' and pho_all ='y' )"
    ms.open sql,z_photo_cn,1,3
	pho_mx_all=ms("pho_mx_all")
	ms.close
	    sql="select count(*) as pho_mx_sh from z_photo_group where (pho_state = 'v' and pho_shared = 'y' and pho_all ='n' )"
    ms.open sql,z_photo_cn,1,3
	pho_mx_sh=ms("pho_mx_sh")
	ms.close	
		
		 %>
        <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="tableBorder1">
          <tr> 
            <th height="25" colspan="2"><div align="left">&gt;&gt; ���ͳ�ƣ��� <%= pho_mx %> ����ᣬ���� 
                <%= pho_mx_all %> ���������, <%= pho_mx_sh %> �����������᣻</div></th>
          </tr>
          <tr> 
            <td width="9%" height="30" class="tableBody2"><div align="center"><font color="red">�������</font></div></td>
            <td width="91%" height="30" class="tableBody1"><table width="98%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			<% 
			dim m_username,m_pho_group_name,m_pho_totle
	sql="select top 5 * from z_photo_group where (pho_state = 'v' and pho_all='n') order by pho_group_id desc"
    ms.open sql,z_photo_cn,1,3
	do while not ms.eof
	m_username=ms("username")
	m_pho_group_name=ms("pho_group_name")
	'm_pho_totle=ms("pho_totle")
	
			 %>
			      <td><a href="z_photo3.asp?pho_names=<%=m_username%>">��<%= m_pho_group_name %>��(<%= m_username %>)</a></td>
		<%  ms.movenext
			loop
			ms.close
			 %>
                </tr>
              </table> </td>
          </tr>
          <tr> 
            <td height="30" class="tableBody2"><div align="center"><font color="red">�������</font></div></td>
            <td height="30" class="tableBody1">
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
			  <tr>			
						<% 
	dim k
	k=0
	sql="select top 4 * from z_photo_group where (pho_state = 'v' and pho_all='n') order by pho_group_totle desc"
    ms.open sql,z_photo_cn,1,3
	do while not ms.eof and k<5
	m_username=ms("username")
	m_pho_group_name=ms("pho_group_name")
	m_pho_totle=ms("pho_group_totle")
	
			 %>
			      <td><a href="z_photo3.asp?pho_names=<%=m_username%>">��<%= m_pho_group_name %>��(<%= m_pho_totle %>)</a></td>
		<%  ms.movenext
		k=k+1
			loop
			ms.close
			 %>
		  </tr>
              </table> </td>
			</tr>
        </table>
		
      </form>
      <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td height="20">&nbsp;</td>
        </tr>
        <tr> 
          <td valign="top"> <div align="center"></div>
            <table width="100%" height="600" border="0" align=center cellpadding="0" cellspacing="0" class=tableborder1>
              <tr> 
                <td valign=top class="tablebody2"> <div align="left"></div>
                  <table width="100%" height="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder1">
                    <tr> 
                      <th height="22">�����������б�</th>
                    </tr>
                    <tr> 
                      <td height="73" class="tablebody2">
						<form name="form1" method="post" action="z_photo2.asp">
                          <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder1">
                            <tr> 
                              <td width="11%" class="forumRowHighlight"><div align="center">�� 
                                  �� ��</div></td>
                              <td width="23%" class="forumRowHighlight"><div align="center">�������</div>
                                <div align="center"></div></td>
                              <td colspan="2" class="forumRowHighlight"><div align="center">�� 
                                  ��</div></td>
                              <td width="17%" class="forumRowHighlight"><div align="center">�� 
                                  ��</div></td>
                            </tr>
                            <% 
dim ps,mx,mp,cp,page,i
dim username,pho_group_name,pho_group_brief,pho_group_totle
dim tmpusername
ps=20

page=request.querystring("page")
if page="" then
cp=1
else
cp=cint(page)
end if

 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where (pho_state = 'v' and pho_shared = 'y' and pho_all ='n' ) order by username desc "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td colspan=4 class=tablebody2>��Ȼ��û��һ����ᣬ��ҿ�ȥ��һ������</td>"
    Response.Write "</tr>"
else
rs.Pagesize=ps
mx=rs.recordcount
mp=rs.pagecount

if mp<cp then
rs.absolutepage=mp
else
rs.absolutepage=cp
end if

i=0
do while not rs.eof and i<ps
username=rs("username")
pho_group_name=rs("pho_group_name")
pho_group_brief=rs("pho_group_brief")
pho_group_totle=rs("pho_group_totle")




rs.movenext
i=i+1
 %>
                            <tr> 
                              <td class="tablebody2"><div align="center"><a href="z_photo3.asp?pho_names=<%=username%>"><%= username %></a></div></td>
                              <td class="tablebody1"><div align="center">��<%= pho_group_name %>��</div></td>
                              <td colspan="2" class="tablebody1"><%= pho_group_brief %></td>
                              <td class="tablebody1"><div align="center"><%= pho_group_totle %></div></td>
                            </tr>
                            <%
loop
end if
rs.close
 %>
                            <tr> 
                              <td class=forumRowHighlight>&nbsp;</td>
                              <td height=21 colspan="2" class=forumRowHighlight>&nbsp; 
                              </td>
                              <td  colspan="2" class="forumRowHighlight"><div align="center"> 
                                  <% if cp>1 then%>
                                  <a href="z_photo4.asp?page=<%=1%>"><font face="Webdings">7</font></a> 
                                  <a href="z_photo4.asp?page=<%=cp-1%>"><font face="Webdings">3</font></a> 
                                  <%end if%>
                                  <%if cp<mp then%>
                                  <a href="z_photo4.asp?page=<%=cp+1%>"><font face="Webdings">4</font></a> 
                                  <a href="z_photo4.asp?page=<%=mp%>"><font face="Webdings">8</font></a> 
                                  <%end if %>
                                </div></td>
                            </tr>
                            <tr> 
                              <td colspan="5" class="tablebody1"><div align="center">��<font color="red"> 
                                  <%= mx %> </font>����<font color="red"><%= mp %> </font>ҳ : 
                                  <% for i=1 to mp %>
                                  <a href="z_photo4.asp?page=<%=i%>">��<%=i%>��</a> 
                                  <% next %>
                                </div></td>
                            </tr>
                          </table>
                        </form></td>
                    </tr>
                    <tr> 
                      <td class="tablebody2">
                       </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <p></p></td>
        </tr>
      </table>
      <p></p></tr>
            </table></td>
        </tr>
      </table>

	</td></tr></table>

<% 

	end sub%>

<% 
'--------------------�����ƶ�-------------------------------------
sub m_move()
	dim abc,ss,a,b,c,d
	d=request.Form("pho_groups")
	for each abc in request.Form("abc")
	ss=split(abc,",")
	a=ss(0)		'old_group_id
	b=ss(1)		'pho_size
	c=ss(2)		'pho_id
	call movepho(a,b,c,d)
	next

	Response.Write "<script>"
	Response.Write "alert('�Ѿ��ɹ��ƶ���');"
	Response.Write "</script>"
end sub
'--------------------����ɾ��--------------------------------------
sub m_del()
	dim abc,ss,a,b,c
	for each abc in request.Form("abc")
	ss=split(abc,",")
	a=ss(0)
	b=ss(1)
	c=ss(2)
	call del(a,b,c)
	next
	Response.Write "<script>"
	Response.Write "alert('�Ѿ��ɹ�ɾ����');"
	Response.Write "</script>"
end sub
'------------------�ƶ�����-------------------------------
sub movepho(a,b,c,d)

 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo where pho_id = "&c
		rs.open sql,z_photo_cn,1,3
		rs("pho_group")=d
		rs.update
		rs.close	
call up_group(d,b)
call down_group(a,b)
end sub
 '--------------------------ɾ����Ƭ---------------------------------
sub del(a,b,c)

 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo where pho_id = "&c
		rs.open sql,z_photo_cn,1,3
		rs("state")="d"
		rs.update
		rs.close	
call down_group(a,b)

end sub 

'-----------------�������ͳ����Ϣ-����-------------------
sub down_group(d,s)

 set rs=server.createobject("adodb.recordset")
    sql="select pho_group_totle,pho_group_size from z_photo_group where pho_group_id = "&d
		rs.open sql,z_photo_cn,1,3
		rs("pho_group_totle")=rs("pho_group_totle")-1
		rs("pho_group_size")=rs("pho_group_size")-s
		rs.update
		rs.close

end sub
'-----------------�������ͳ����Ϣ-����-------------------
sub up_group(d,s)

d=cint(d)
s=cint(s)
 set rs=server.createobject("adodb.recordset")
    sql="select pho_group_totle,pho_group_size from z_photo_group where pho_group_id = "&d
		rs.open sql,z_photo_cn,1,3
		rs("pho_group_totle")=rs("pho_group_totle")+1
		rs("pho_group_size")=rs("pho_group_size")+s
		rs.update
		rs.close
end sub
%>