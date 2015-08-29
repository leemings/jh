<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<script language="JavaScript" type="text/JavaScript">
function z_photo_checkform()
{
if(document.all["pho_group_name"].value==""){
	alert("相册名称不能为空！");
	return false;
}
if(document.all["pho_group_brief"].value==""){
	alert("相册说明不能为空！");
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
	form1.add.value=" 修 改 ";
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
	form1.add.value=" 修 改 ";
	form1.action="z_photo_group.asp?menu=2";
}

function add1(){
	form1.reset();
	form1.add.value=" 添 加 ";
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
			'call del()				'默认
		case 1
			call newgroup()			'添加新的相册组
		case 2
			call editgroup()		'修改相册组
		case 3
			call delgroup()			'删除相册组
		case 4
			call countgroup()		'更新相册组统计信息
		case else 
														
	end select


stats="个人相册"

call nav()

call head_var(2,0,"","")

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入个人相册的权限，请先登录或者同管理员联系。"
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
		
    <th height=25>个 人 相 册</td> </tr>
	<tr>
		<td class=tablebody2 height=1>	

<br>
      <form action="z_photo3.asp" method="get" name="form2" id="form2">
        <table height="25" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tr> 
            <td align=left valign=middle class=TableTitle2><font face="Webdings">2</font>&nbsp;<font color="BLUE"><a href="#edit1" onclick="add1();">添加相册组</a></font> 
             <font face="Webdings">2</font>&nbsp;<font color=blue><a href="z_photo.asp">相册首页</a></font> 
              <font face="Webdings">2</font>&nbsp;<a href="z_photo2.asp">照片管理</a><%if master then %>&nbsp;<font face="Webdings">2</font>&nbsp;<a href="z_photo_admin.asp">总管理</a><% end if %></td>
            <td align=right valign=middle class=TableTitle2>用户名： 
              <input name="pho_names" type="text" id="pho_names" size="16"> <input type="submit" name="Submit" value="浏览他的展示相册"> 
            </td>
          </tr>
        </table>
      </form>

	  <br>
      <table width="89%" border="0" align="center" cellpadding="3" cellspacing="1" class=tableborder1>
        <tr> 
          <th height="20" colspan="6"><div align="left">-=&gt; 相册组管理</div></th>
        </tr>
        <tr> 
          <td height="25" colspan="6" class=tablebody2><div align="center"><font color="red">『公共相册』</font></div></td>
        </tr>
        <tr> 
          <td width="15%" height="20" class=TopLighNav1><div id=p1 align="center"><strong>名 
              称</strong></div></td>
          <td width="46%" height="20" class=TopLighNav1><div id=p2 align="center"><strong>说 
              明</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div id=p3 align="center"><strong>数 
              量</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div align="center"><strong>空 
              间</strong></div></td>
          <td width="5%" class=TopLighNav1><div id=p4 align="center"><strong>属 
              性</strong></div></td>
          <td width="15%" height="20" class=TopLighNav1><div id=p5 align="center"><strong>管 
              理</strong></div></td>
        </tr>
        <% 
dim per_id,per_name,per_brief,per_totle,per_all,per_shared,i,pshare,del_id,per_size
i=0
set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where pho_all='y' and pho_state='v' "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td colspan='5' class=tablebody1>目前还没有相册！</td>"
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
pshare="保密"
else
pshare="共享"
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
          <td class=tablebody1 ><div id=p5 value="<%=per_all%>" " align="center"><a href="z_photo_group.asp?del_id=<%=per_id%>&menu=4">更新</a> 
              <a href="#edit1" onclick="edit_admin(<%=i%>);" >编辑</a> 
              <a href="z_photo_group.asp?del_id=<%=per_id%>&menu=3">删除</a></div></td>
          <% end if %>
        </tr>
        <% 
loop
end if
rs.close
%>
        <tr> 
          <td height="25" colspan="6" class=tablebody2><div align="center"><font color="red">『您的相册』</font></div></td>
        </tr>
        <tr> 
          <td width="15%" height="20" class=TopLighNav1><div align="center"><strong>名 
              称</strong></div></td>
          <td width="50%" height="20" class=TopLighNav1><div align="center"><strong>说 
              明</strong></div></td>
          <td width="2%" height="20" class=TopLighNav1><div align="center"><strong>数 
              量</strong></div></td>
          <td width="8%" height="20" class=TopLighNav1><div align="center"><strong>空 
              间</strong></div></td>
          <td width="5%" class=TopLighNav1><div align="center"><strong>属 性</strong></div></td>
          <td width="15%" height="20" class=TopLighNav1><div align="center"><strong>管 
              理</strong></div></td>
        </tr>
        <% 


set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where (pho_state = 'v' and username ='"&membername&"' and pho_all='n') "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td colspan='5' class=tablebody1>目前还没有相册！</td>"
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
pshare="保密"
else
pshare="共享"
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
          <td class=tablebody1><div id=p5 value="<%=per_all%>" " align="center"><a href="z_photo_group.asp?del_id=<%=per_id%>&menu=4">更新</a> 
              <a href="#edit1" onclick="edit(<%=i%>);" >编辑</a> 
              <a href="z_photo_group.asp?del_id=<%=per_id%>&menu=3">删除</a></div></td>
        </tr>
        <% 

loop
end if
rs.close%>
        <tr> 
          <td colspan="6" class=TopLighNav1><div align="center">&lt;&gt; 添加/修改相册组 
              &lt;&gt; </div></td>
        </tr>
        <tr> 
          <td colspan="6" class=tablebody2><div align="center"> 
              <form name="form1" method="post" action="z_photo_group.asp?menu=1" onSubmit="return z_photo_checkform()">
                <table width="90%" border="0" cellspacing="0" cellpadding="3">
                  <tr> 
                    <td><a name="edit1"></a>名 称： 
                      <input name="pho_group_name" type="text" id="pho_group_name" maxlength="20">
                      </td>
                  </tr>
                  <tr> 
                    <td> 设 置： <% if master then  %>
                       <input name="pho_all" type="checkbox" id="pho_all" value="y">
                      是否论坛公共组  <% end if %>
                      <input name="pho_shared" type="checkbox" id="pho_shared" value="y">
                      是否共享（其他用户可见）</td>
                  </tr>
                  <tr> 
                    <td>说 明： 
                      <input name="pho_group_brief" type="text" id="pho_group_brief" size="50" maxlength="50"> 
                      <input name="edit_id" type="hidden" id="edit_id"> <input type="submit" name="add" value=" 添 加 "></td>
                  </tr>
                  <tr> 
                    <td>* 名称字数在20个字以内，说明字数在50个字以内；</td>
                  </tr>
                </table>
              </form>
            </div></td>
        </tr>
      </table></td></tr></table>

<% 

	end sub%>

<% 
'---------公共程序区--------------
'------------添加新的相册组--------------------
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
Response.Write "alert('您已成功添加一个相册！');"
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
Response.Write "alert('您已成功修改该相册！');"
Response.Write "</script>"
end sub
'--------------删除相册组----------------------
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
Response.Write "alert('已经成功删除该相册！');"
Response.Write "</script>"

end sub
'-----------------更新相册统计信息-------------------
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
Response.Write "alert('已经成功更新该相册统计信息！');"
Response.Write "</script>"

end sub
 %>