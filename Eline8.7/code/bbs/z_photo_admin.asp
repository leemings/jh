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
function configset() {
	form3.z_photo_totle_size.value="20480000";
	form3.z_photo_size.value="1024000";
	form3.z_photo_type.value="jpg,gif,jpeg,bmp";
}

function phoidall(form)  {
  for (var i=0;i<form.abc.length;i++)    {
    var e = form.abc[i];
    if (e.name != 'pho_id_all') e.checked = form.pho_id_all.checked; 
	}
}
</script>
<%
dim menu

menu=request.QueryString("menu")
if menu="" then menu=0
stats="�������"
call nav()
call head_var(2,0,"","")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н����������Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	if not master then
		errmsg=errmsg+"<br><li>��û�й����������Ȩ�ޣ���ȷ�����Ѿ���¼���Ҿ�����Ӧ��Ȩ�ޣ�"
		stats="������� ������Ϣ"
		call nav()
		call head_var(0,0,"�������","Z_test.asp")		
		call dvbbs_error()	
	else
		call start()
	end if
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
			call m_move()				'ɾ����Ƭ
		case 3
			call clearall()				'�������ռ�
		case 4
			call setconfig()			'�����������

		case else 
	end select 
'-------------------------
%>
<table width="97%" align=center cellpadding=3 cellspacing=1 class=tableborder1>
	<tr><th height=25>�� �� �� ��</th></tr>
	<tr>
		<td class=tablebody2 height=1><br>
			<form name="form1" method="get" action="z_photo3.asp">
        <table width="98%" height="25" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tr> 
            <td align=left valign=middle class=TableTitle2><font face="Webdings">2</font>&nbsp;<font color="BLUE"><a href="z_photo.asp">�����ҳ</a></font> <font face="Webdings">2</font><a href="z_photo4.asp"><font color="blue">&nbsp;</font>����б�</a> <font face="Webdings">2</font>&nbsp;<font color=blue><a href="z_photo_group.asp">��������</a></font> <font face="Webdings">2</font>&nbsp;<a href="z_photo2.asp">��Ƭ����</a> <%if master then %>&nbsp;<font face="Webdings">2</font>&nbsp;<a href="z_photo_admin.asp">�ܹ���</a><% end if %></td>
            <td align=right valign=middle class=TableTitle2>�û����� <input name="pho_names" type="text" id="pho_names" size="16"> <input type="submit" name="Submit" value="����������"></td>
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
            <th height="25" colspan="2"><div align="left">&gt;&gt; ���ͳ�ƣ��� <%= pho_mx %> ����ᣬ���� <%= pho_mx_all %> ���������, <%= pho_mx_sh %> �����������᣻</div></th>
          </tr>
          <tr> 
            <td width="9%" height="30" class="tableBody2"><div align="center"><font color="red">�������</font></div></td>
            <td width="91%" height="30" class="tableBody1">
            	<table width="98%" border="0" cellspacing="0" cellpadding="0">
			  				<tr><% 
									dim m_username,m_pho_group_name,m_pho_totle
									sql="select top 5 * from z_photo_group where (pho_state = 'v' and pho_all='n') order by pho_group_id desc"
									ms.open sql,z_photo_cn,1,3
									do while not ms.eof
										m_username=ms("username")
										m_pho_group_name=ms("pho_group_name")
										'm_pho_totle=ms("pho_totle")
										%><td><a href="z_photo3.asp?pho_names=<%=m_username%>">��<%= m_pho_group_name %>��(<%= m_username %>)</a></td><%
										ms.movenext
									loop
									ms.close
								%></tr>
							</table>
						</td>
					</tr>
					<tr> 
          	<td height="30" class="tableBody2"><div align="center"><font color="red">�������</font></div></td>
          	<td height="30" class="tableBody1">
          		<table width="98%" border="0" cellspacing="0" cellpadding="0">
								<tr><% 
									dim k
									k=0
									sql="select top 4 * from z_photo_group where (pho_state = 'v' and pho_all='n') order by pho_group_totle desc"
									ms.open sql,z_photo_cn,1,3
									do while not ms.eof and k<5
										m_username=ms("username")
										m_pho_group_name=ms("pho_group_name")
										m_pho_totle=ms("pho_group_totle")
										%><td><a href="z_photo3.asp?pho_names=<%=m_username%>">��<%= m_pho_group_name %>��(<%= m_pho_totle %>)</a></td><%
										ms.movenext
										k=k+1
									loop
									ms.close
								%></tr>
							</table>
						</td>
					</tr>
        </table>
      </form>
      <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr><td height="20">&nbsp;</td></tr>
        <tr> 
          <td valign="top">
            <table width="100%" height="600" border="0" align=center cellpadding="0" cellspacing="0" class=tableborder1>
              <tr> 
                <td valign=top class="tablebody2">
                  <table width="100%" height="100%" border="0" cellpadding="3" cellspacing="1" class="tableborder1">
                    <tr><th height="22">��������</th></tr>
                    <tr> 
                      <td height="73" class="tablebody2"><form name="form1" method="post" action="z_photo2.asp">
												<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder1">
													<tr> 
                            <td width="4%" class="forumRow">&nbsp;</td>
                            <td width="63%" class="forumRow"><div align="center">�ꡡ��</div></td>
                            <td width="16%" class="forumRow"><div align="center">��С</div></td>
                            <td width="17%" class="forumRow"><div align="center">ʱ����</div></td>
                          </tr><% 
													dim ps,mx,mp,cp,page,i
													dim pho_url1,pho_title1,pho_brief1,pho_adtime1,pho_id1,username1,pho_count1,pho_group1,pho_size1
													dim abc
													ps=20
													
													page=request.querystring("page")
													if page="" then
														cp=1
													else
														cp=cint(page)
													end if
													
													set rs=server.createobject("adodb.recordset")
													sql="select * from z_photo where (state = 'v' ) order by pho_id desc "
													rs.open sql,z_photo_cn,1,3
													if rs.eof then
														Response.Write "<tr>" 
														Response.Write "<td colspan=4 class=tablebody2>û����Ƭ��</td>"
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
															pho_id1=rs("pho_id")
															'username1=rs("username")
															pho_url1=rs("pho_url")
															pho_adtime1=rs("pho_adtime")
															pho_adtime1=formatdatetime(pho_adtime1,2)
															pho_title1=rs("pho_title")
															pho_group1=rs("pho_group")
															pho_size1=rs("pho_size")
															abc=pho_group1&","&pho_size1&","&pho_id1
															rs.movenext
															i=i+1
															%><tr> 
																<td class="tablebody2"><div align="center"><input name="abc" type="checkbox" id="abc" value="<%= abc %>"></div></td>
																<td class="tablebody1"><a href="javascript:openScript('z_photo_view1.asp?pho_id=<%=pho_id1%>&action=new',550,590)"><%= pho_title1 %></a></td>
																<td class="tablebody1"><div align="center"><%= pho_size1 %></div></td>
																<td class="tablebody1"><div align="center"><%= pho_adtime1 %></div></td>
															</tr><%
														loop
													end if
													rs.close
													%><tr> 
														<td class=forumRow>&nbsp;</td>
														<td height=21 class=forumRow> ȫ�� <input name="pho_id_all" type="checkbox" id="pho_id_all2" value="checkbox" onpropertychange="phoidall(this.form)"> ת���� <select name="pho_groups" id="pho_groups"><% 
															dim pho_group_id,per_name
															set rs=server.createobject("adodb.recordset")
															sql="select * from z_photo_group where (pho_state = 'v' and pho_all='y') "
															rs.open sql,z_photo_cn,1,3
															do while not rs.eof 
																pho_group_id=rs("pho_group_id")
																per_name=rs("pho_group_name")
																%><option selected value="<%= pho_group_id %>" ><%= per_name %></option><%
																rs.movenext 
															loop
															rs.close
															set rs=server.createobject("adodb.recordset")
															sql="select * from z_photo_group where (pho_state = 'v' and username='"&membername&"' and pho_all='n') "
															rs.open sql,z_photo_cn,1,3
															if not rs.eof then
																do while not rs.eof 
																	pho_group_id=rs("pho_group_id")
																	per_name=rs("pho_group_name")
																	%><option value="<%= pho_group_id %>" ><%= per_name %></option><%
																	rs.movenext 
																loop
																rs.close
																%></select><% 
															else 
																Response.Write "</select>"
																Response.Write "�㻹û�н���������ᣬ�뵽��������ҳ������ӣ�"
															end if
															%> <input type="radio" name="menu" value="2"> ת�� <input type="radio" name="menu" value="1"> ɾ�� <input name="go" type="submit" id="go2" value="ִ��">
														</td>
														<td  colspan="2" class="forumRow"><div align="center"> <% if cp>1 then%><a href="z_photo_admin.asp?page=<%=1%>"><font face="Webdings">7</font></a> <a href="z_photo_admin.asp?page=<%=cp-1%>"><font face="Webdings">3</font></a> <%end if%><%if cp<mp then%><a href="z_photo_admin.asp?page=<%=cp+1%>"><font face="Webdings">4</font></a> <a href="z_photo_admin.asp?page=<%=mp%>"><font face="Webdings">8</font></a> <%end if %></div></td>
													</tr>
													<tr> 
														<td colspan="4" class="tablebody1"><div align="center">��<font color="red"> <%= mx %> </font>�ţ�<font color="red"><%= mp %> </font>ҳ : <% for i=1 to mp %><a href="z_photo_admin.asp?page=<%=i%>">��<%=i%>��</a> <% next %></div></td>
													</tr>
												</table>
											</form></td>
                    </tr>
                    <tr>
                      <td class="tablebody2"><div align="center"><%
												dim dx,ds,tx,ts
												set rs=server.createobject("adodb.recordset")
												sql="select count(*) as tx,sum(pho_size) as ts from z_photo "
												rs.open sql,z_photo_cn,1,3
												tx=rs("tx")
												ts=rs("ts")
												rs.close
												set rs=server.createobject("adodb.recordset")
												sql="select count(*) as dx,sum(pho_size) as ds from z_photo where state = 'd' "
												rs.open sql,z_photo_cn,1,3
												dx=rs("dx")
												ds=rs("ds")
												rs.close
												%></div>
                        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder1">
													<tr><th height="22" colspan="4">���ϵͳ�ռ����</th></tr>
                          <tr> 
														<td width="15%" class="tablebody1"><div align="right">Ŀǰʹ�ÿռ䣺</div></td>
                            <td width="36%" class="tablebody1"><%= ts %></td>
                            <td width="15%" class="tablebody1"><div align="right">Ŀǰ���м�¼��</div></td>
                            <td width="34%" class="tablebody1"><%= tx %></td>
                          </tr>
                          <tr> 
                            <td class="tablebody1"><div align="right">���ͷſռ䣺</div></td>
                            <td class="tablebody1"><%= ds %></td>
                            <td class="tablebody1"><div align="right">���ͷż�¼��</div></td>
                            <td class="tablebody1"><%= dx %></td>
                          </tr>
                          <tr> 
                            <td colspan="4" class="tablebody1"><form name="form2" method="post" action="z_photo_admin.asp?menu=3"><div align="center"><input type="submit" name="Submit" value="�����ջؿ��ͷſռ�"> <font color="red"> �˲�����ɾ����Ӳ���е��Ѿ���Ϊ���ʹ�õ��ļ� </font></div></form></td>
                          </tr>
                        </table>
											</td>
                    </tr>
										<tr><% 
											dim z_photo_totle_size,z_photo_size,z_photo_type
											set rs=server.createobject("adodb.recordset")
											sql="select * from z_photo_config where id=1"
											rs.open sql,z_photo_cn,1,3
											z_photo_totle_size=rs("z_photo_totle_size")
											z_photo_size=rs("z_photo_size")
											z_photo_type=rs("z_photo_type")
											rs.close
											%><td class="tablebody2">
												<form name="form3" method="post" action="z_photo_admin.asp?menu=4">
													<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder1">
														<tr><th height="22" colspan="3">���ϵͳ��������</th></tr>
                            <tr>
                              <td width="15%" class="tablebody1"><div align="right">ÿ���û��ռ䣺</div></td>
                              <td width="83%" colspan="2" class="tablebody1"><input name="z_photo_totle_size" type="text" id="z_photo_totle_size" value="<%= z_photo_totle_size %>"> <font color="red">Ĭ��Ϊ��1024*20000��1024=1K;</font></td>
                            </tr>
                            <tr> 
                              <td class="tablebody1"><div align="right">������Ƭ���ƣ�</div></td>
                              <td colspan="2" class="tablebody1"><input name="z_photo_size" type="text" id="z_photo_size" value="<%= z_photo_size %>"> <font color="red">Ĭ��Ϊ��1024*1000��1024=1K;</font></td>
                            </tr>
                            <tr> 
                              <td class="tablebody1"><div align="right">�ϴ��ļ����ͣ�</div></td>
                              <td colspan="2" class="tablebody1"><input name="z_photo_type" type="text" id="z_photo_type" value="<%= z_photo_type %>"> <font color="red">Ĭ��Ϊ��jpg,gif,jpeg</font></td>
                            </tr>
                            <tr> 
                              <td class="tablebody1">&nbsp;</td>
                              <td colspan="2" class="tablebody1"><input type="submit" name="Submit2" value="����/Ӧ��"><input type="button" name="Submit3" value="Ĭ������" onclick="configset()"> <input type="reset" name="Submit4" value="ȡ��"></td>
														</tr>
                          </table>
												</form>
											</td>
										</tr>
                  </table>
								</td>
              </tr>
            </table>
					</td>
        </tr>
      </table>
		</td>
	</tr>
</table>
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

d=cint(d)
s=cint(s)
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

sub clearall()
	dim d_pho_url,fsd
	set fsd=createobject("Scripting.filesystemobject")
	 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo where state = 'd' "
    rs.open sql,z_photo_cn,1,3
	if not rs.eof then
	do while not rs.eof
	d_pho_url=rs("pho_url")
	if(left(d_pho_url,7)="photos/") then
	d_pho_url=Server.mappath(d_pho_url)
	if (fsd.FileExists(d_pho_url)=true) then
	fsd.DeleteFile(d_pho_url)
	end if
	end if
	rs.delete
	rs.update
	rs.movenext
	loop
	end if
	rs.close
	Response.Write "<script>"
	Response.Write "alert('�Ѿ��ɹ��������ϵͳ��');"
	Response.Write "</script>"
end sub

sub setconfig()
	dim a,b,c

	a=request("z_photo_totle_size")
	b=request("z_photo_size")
	c=request("z_photo_type")
	z_photo_cn.execute("update z_photo_config set z_photo_totle_size="&a&",z_photo_size="&b&",z_photo_type='"&c&"'")
	Response.Write "<script>"
	Response.Write "alert('�Ѿ��ɹ��޸����ϵͳ���ã�');"
	Response.Write "</script>"
end sub

%>