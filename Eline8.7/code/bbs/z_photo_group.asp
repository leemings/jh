<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<script language="JavaScript" type="text/JavaScript">
function z_photo_checkform()
{
if(document.all["pho_group_name"].value==""){
	alert("������Ʋ���Ϊ�գ�");
	return false;
}
if(document.all["pho_group_brief"].value==""){
	alert("���˵������Ϊ�գ�");
	return false;
}
return true;

}

function edit_admin(n){
//alert(n);
	form1.pho_group_name.value=document.all["p1"][n].innerHTML;
	form1.pho_group_brief.value=document.all["p2"][n].innerHTML;
	form1.edit_id.value=document.all["p3"][n].value;
	if(document.all["p5"][n].value=="y")
	form1.pho_all.checked=true;
else {
	form1.pho_all.checked=false;
}
	if(document.all["p4"][n].value=="y")
	form1.pho_shared.checked=true;
else {
	form1.pho_shared.checked=false;
}
	form1.add.value=" �� �� ";
	form1.action="z_photo_group.asp?menu=2";
}
function edit(n){
//alert(n);
	form1.pho_group_name.value=document.all["p1"][n].innerHTML;
	form1.pho_group_brief.value=document.all["p2"][n].innerHTML;
	form1.edit_id.value=document.all["p3"][n].value;

	if(document.all["p4"][n].value=="y")
	form1.pho_shared.checked=true;
else {
	form1.pho_shared.checked=false;
}
	form1.add.value=" �� �� ";
	form1.action="z_photo_group.asp?menu=2";
}

function add1(){
	form1.reset();
	form1.add.value=" �� �� ";
	form1.action="z_photo_group.asp?menu=1";
}
</script>
<%
dim menu
menu=request.QueryString("menu")
if menu="" then menu=0
'Response.Write(menu)
select case menu
		case 0
			'call del()				'Ĭ��
		case 1
			call newgroup()			'����µ������
		case 2
			call editgroup()		'�޸������
		case 3
			call delgroup()			'ɾ�������
		case 4
			call countgroup()		'���������ͳ����Ϣ
		case else 
														
	end select


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
  

'-------------------------
%><title></title>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr>
		
    <th height=25>�� �� �� ��</td> </tr>
	<tr>
		<td class=tablebody2 height=1>	

<br>
      <form action="z_photo3.asp" method="get" name="form2" id="form2">
        <table height="25" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tr> 
            <td align=left valign=middle class=TableTitle2><font face="Webdings">2</font>&nbsp;<font color="BLUE"><a href="#edit1" onclick="add1();">��������</a></font> 
             <font face="Webdings">2</font>&nbsp;<font color=blue><a href="z_photo.asp">�����ҳ</a></font> 
              <font face="Webdings">2</font>&nbsp;<a href="z_photo2.asp">��Ƭ����</a><%if master then %>&nbsp;<font face="Webdings">2</font>&nbsp;<a href="z_photo_admin.asp">�ܹ���</a><% end if %></td>
            <td align=right valign=middle class=TableTitle2>�û����� 
              <input name="pho_names" type="text" id="pho_names" size="16"> <input type="submit" name="Submit" value="�������չʾ���"> 
            </td>
          </tr>
        </table>
      </form>

	  <br>
      <table width="89%" border="0" align="center" cellpadding="3" cellspacing="1" class=tableborder1>
        <tr> 
          <th height="20" colspan="6"><div align="left">-=&gt; ��������</div></th>
        </tr>
        <tr> 
          <td height="25" colspan="6" class=tablebody2><div align="center"><font color="red">��������᡻</font></div></td>
        </tr>
        <tr> 
          <td width="15%" height="20" class=TopLighNav1><div id=p1 align="center"><strong>�� 
              ��</strong></div></td>
          <td width="46%" height="20" class=TopLighNav1><div id=p2 align="center"><strong>˵ 
              ��</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div id=p3 align="center"><strong>�� 
              ��</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div align="center"><strong>�� 
              ��</strong></div></td>
          <td width="5%" class=TopLighNav1><div id=p4 align="center"><strong>�� 
              ��</strong></div></td>
          <td width="15%" height="20" class=TopLighNav1><div id=p5 align="center"><strong>�� 
              ��</strong></div></td>
        </tr>
        <% 
dim per_id,per_name,per_brief,per_totle,per_all,per_shared,i,pshare,del_id,per_size
i=0
set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where pho_all='y' and pho_state='v' "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td colspan='5' class=tablebody1>Ŀǰ��û����ᣡ</td>"
    Response.Write "</tr>"
else
do while not rs.eof 
per_id=rs("pho_group_id")
per_name=rs("pho_group_name")
per_brief=rs("pho_group_brief")
per_totle=rs("pho_group_totle")
per_all=rs("pho_all")
per_shared=rs("pho_shared")
per_size=rs("pho_group_size")
if per_shared="n" then
pshare="����"
else
pshare="����"
end if 
rs.movenext
i=i+1
 %>
        <tr> 
          <td class=tablebody1><div id=p1 align="center"><%= per_name %></div></td>
          <td class=tablebody1><div id=p2 align="center"><%= per_brief %></div></td>
          <td class=tablebody1><div id=p3 align="center" value="<%= per_id %>"><%= per_totle %></div></td>
          <td class=tablebody1><div align="center"><%= per_size %></div></td>
          <td class=tablebody1><div id=p4 align="center" value="<%= per_shared %>"><%= pshare %></div></td>
          <% if master then %>
          <td class=tablebody1 ><div id=p5 value="<%=per_all%>" " align="center"><a href="z_photo_group.asp?del_id=<%=per_id%>&menu=4">����</a> 
              <a href="#edit1" onclick="edit_admin(<%=i%>);" >�༭</a> 
              <a href="z_photo_group.asp?del_id=<%=per_id%>&menu=3">ɾ��</a></div></td>
          <% end if %>
        </tr>
        <% 
loop
end if
rs.close
%>
        <tr> 
          <td height="25" colspan="6" class=tablebody2><div align="center"><font color="red">��������᡻</font></div></td>
        </tr>
        <tr> 
          <td width="15%" height="20" class=TopLighNav1><div align="center"><strong>�� 
              ��</strong></div></td>
          <td width="50%" height="20" class=TopLighNav1><div align="center"><strong>˵ 
              ��</strong></div></td>
          <td width="2%" height="20" class=TopLighNav1><div align="center"><strong>�� 
              ��</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div align="center"><strong>�� 
              ��</strong></div></td>
          <td width="5%" class=TopLighNav1><div align="center"><strong>�� ��</strong></div></td>
          <td width="15%" height="20" class=TopLighNav1><div align="center"><strong>�� 
              ��</strong></div></td>
        </tr>
        <% 


set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where (pho_state = 'v' and username ='"&membername&"' and pho_all='n') "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td colspan='5' class=tablebody1>Ŀǰ��û����ᣡ</td>"
    Response.Write "</tr>"
else
do while not rs.eof 
per_id=rs("pho_group_id")
per_name=rs("pho_group_name")
per_brief=rs("pho_group_brief")
per_totle=rs("pho_group_totle")
per_all=rs("pho_all")
per_shared=rs("pho_shared")
per_size=rs("pho_group_size")
if per_shared="n" then
pshare="����"
else
pshare="����"
end if 
rs.movenext
i=i+1
 %>
        <tr> 
          <td class=tablebody1><div id=p1 align="center"><%= per_name %></div></td>
          <td class=tablebody1><div id=p2 align="center"><%= per_brief %></div></td>
          <td class=tablebody1><div id=p3 align="center" value="<%= per_id %>"><%= per_totle %></div></td>
          <td class=tablebody1><div align="center"><%= per_size %></div></td>
          <td class=tablebody1><div id=p4 align="center" value="<%= per_shared %>"><%= pshare %></div></td>
          <td class=tablebody1><div id=p5 value="<%=per_all%>" " align="center"><a href="z_photo_group.asp?del_id=<%=per_id%>&menu=4">����</a> 
              <a href="#edit1" onclick="edit(<%=i%>);" >�༭</a> 
              <a href="z_photo_group.asp?del_id=<%=per_id%>&menu=3">ɾ��</a></div></td>
        </tr>
        <% 

loop
end if
rs.close%>
        <tr> 
          <td colspan="6" class=TopLighNav1><div align="center">&lt;&gt; ���/�޸������ 
              &lt;&gt; </div></td>
        </tr>
        <tr> 
          <td colspan="6" class=tablebody2><div align="center"> 
              <form name="form1" method="post" action="z_photo_group.asp?menu=1" onSubmit="return z_photo_checkform()">
                <table width="90%" border="0" cellspacing="0" cellpadding="3">
                  <tr> 
                    <td><a name="edit1"></a>�� �ƣ� 
                      <input name="pho_group_name" type="text" id="pho_group_name" maxlength="20">
                      ��</td>
                  </tr>
                  <tr> 
                    <td> �� �ã� <% if master then  %>
                       <input name="pho_all" type="checkbox" id="pho_all" value="y">
                      �Ƿ���̳������ �� <% end if %>
                      <input name="pho_shared" type="checkbox" id="pho_shared" value="y">
                      �Ƿ��������û��ɼ���</td>
                  </tr>
                  <tr> 
                    <td>˵ ���� 
                      <input name="pho_group_brief" type="text" id="pho_group_brief" size="50" maxlength="50"> 
                      <input name="edit_id" type="hidden" id="edit_id"> <input type="submit" name="add" value=" �� �� "></td>
                  </tr>
                  <tr> 
                    <td>* ����������20�������ڣ�˵��������50�������ڣ�</td>
                  </tr>
                </table>
              </form>
            </div></td>
        </tr>
      </table></td></tr></table>

<% 

	end sub%>

<% 
'---------����������--------------
'------------����µ������--------------------
sub newgroup()
dim pho_group_name,pho_group_brief,pho_shared,pho_all
pho_group_name=request.form("pho_group_name")
pho_group_brief=request.form("pho_group_brief")
pho_all=request.form("pho_all")
if pho_all="" then pho_all="n"
pho_shared=request.form("pho_shared")
if pho_shared="" then pho_shared="n"
z_photo_cn.execute("insert into z_photo_group (username,pho_group_name,pho_group_brief,pho_all,pho_shared) values ('"&membername&"','"&pho_group_name&"','"&pho_group_brief&"','"&pho_all&"','"&pho_shared&"')")
Response.Write "<script>"
Response.Write "alert('���ѳɹ����һ����ᣡ');"
Response.Write "</script>"
end sub

sub editgroup()
dim pho_group_name,pho_group_brief,pho_shared,pho_all,edit_id
edit_id=cint(request.form("edit_id"))
pho_group_name=request.form("pho_group_name")
pho_group_brief=request.form("pho_group_brief")
pho_all=request.form("pho_all")
if pho_all="" then pho_all="n"
pho_shared=request.form("pho_shared")
if pho_shared="" then pho_shared="n"

 set rs=server.createobject("adodb.recordset")
   sql="select * from z_photo_group where pho_group_id = "&edit_id
		rs.open sql,z_photo_cn,1,3
		rs("pho_group_name")=pho_group_name
		rs("pho_group_brief")=pho_group_brief
		rs("pho_all")=pho_all
		rs("pho_shared")=pho_shared
		rs.update
		rs.close

Response.Write "<script>"
Response.Write "alert('���ѳɹ��޸ĸ���ᣡ');"
Response.Write "</script>"
end sub
'--------------ɾ�������----------------------
sub delgroup()
dim del_id
del_id=request.QueryString("del_id")
 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where pho_group_id = "&del_id
		rs.open sql,z_photo_cn,1,3
		rs("pho_state")="d"
		rs.update
		rs.close	

Response.Write "<script>"
Response.Write "alert('�Ѿ��ɹ�ɾ������ᣡ');"
Response.Write "</script>"

end sub
'-----------------�������ͳ����Ϣ-------------------
sub countgroup()
dim del_id,mx,sx
del_id=request.QueryString("del_id")
del_id=cint(del_id)
 set rs=server.createobject("adodb.recordset")
    sql="select count(*) as mx,sum(pho_size) as sx from z_photo where state='v' and pho_group = '"&del_id&"'"
		rs.open sql,z_photo_cn,1,3
if not rs.eof then
		mx=rs("mx")
		sx=rs("sx")
else
mx=3
sx=3
end if
if len(sx)>0 then
else
sx=0
end if
rs.close
z_photo_cn.execute("update z_photo_group set pho_group_totle="&mx&",pho_group_size="&sx&" where pho_group_id="&del_id)
			

Response.Write "<script>"
Response.Write "alert('�Ѿ��ɹ����¸����ͳ����Ϣ��');"
Response.Write "</script>"

end sub
 %>