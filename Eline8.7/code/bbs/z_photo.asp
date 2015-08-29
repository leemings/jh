<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->

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

</script>

<%
dim menu

menu=request.querystring("menu")
if menu="" then menu=0

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
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr>
		
    <th height=25>个 人 相 册</td> </tr>
	<tr>
		<td class=tablebody2 height=1>	
<iframe id="pho_if" height="0" frameborder="0"></iframe>
<br>
      <form name="form1" method="get" action="z_photo3.asp">
              <table width="98%" height="25" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tr> 
            <td align=left valign=middle class=TableTitle2><font face="Webdings">2</font>&nbsp;<font color="BLUE"><a href="JavaScript:openScript('z_photo_upload1.asp?action=new',500,400)">添加新照片</a></font> 
              <font face="Webdings">2</font><a href="z_photo4.asp"><font color="blue">&nbsp;</font>相册列表</a> 
              <font face="Webdings">2</font>&nbsp;<font color=blue><a href="z_photo_group.asp">相册组管理</a></font> 
              <font face="Webdings">2</font>&nbsp;<a href="z_photo2.asp">照片管理</a> 
              <%if master then %>&nbsp;<font face="Webdings">2</font>&nbsp;<a href="z_photo_admin.asp">总管理</a><% end if %></td>
            <td align=right valign=middle class=TableTitle2>用户名： 
              <input name="pho_names" type="text" id="pho_names" size="16"> <input type="submit" name="Submit" value="浏览他的相册"> 
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
            <th height="25" colspan="2"><div align="left">&gt;&gt; 相册统计：总 <%= pho_mx %> 个相册，其中 
                <%= pho_mx_all %> 个公共相册, <%= pho_mx_sh %> 个共享个人相册；</div></th>
          </tr>
          <tr> 
            <td width="9%" height="30" class="tableBody2"><div align="center"><font color="red">最新相册</font></div></td>
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
			      <td><a href="z_photo3.asp?pho_names=<%=m_username%>">《<%= m_pho_group_name %>》(<%= m_username %>)</a></td>
		<%  ms.movenext
			loop
			ms.close
			 %>
                </tr>
              </table> </td>
          </tr>
          <tr> 
            <td height="30" class="tableBody2"><div align="center"><font color="red">相册排行</font></div></td>
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
			      <td><a href="z_photo3.asp?pho_names=<%=m_username%>">《<%= m_pho_group_name %>》(<%= m_pho_totle %>)</a></td>
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
      <br>
      <table width="100%" border="0" align=center cellpadding="0" cellspacing="0" class=tableborder1>
        <tr> 
          <td height="22" id=p_group align=right valign=middle class=TopLighNav1> <div align="center"><font id=f_flag color=red>公共相册区</font></div></td>
        </tr>
        <tr >
          <td  class=tablebody2><br>
<% 
dim per_id,per_name,per_brief,per_totle,i
i=0
set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where pho_all='y' and pho_state='v' order by pho_group_id desc"
    rs.open sql,z_photo_cn,1,3
if rs.eof then

    Response.Write " <font color=red>本论坛相册还没有建立一个公共相册，请管理员创立！</font>"

else
do while not rs.eof 
per_id=rs("pho_group_id")
per_name=rs("pho_group_name")
per_brief=rs("pho_group_brief")
per_totle=rs("pho_group_totle")
 
rs.movenext
i=i+1
 %>
            <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder1>
              

				<tr onClick="show(<%= i %>)" style="cursor:hand">
           <th height="25" width="97%"><div align="left">-=&gt; <%= per_name %> <font id=f_flag face="Webdings">6</font></div></th>
              </tr>

				<tr>
         <td height="20" class=tablebody2><% if per_totle<1 then %>该相册中还没有图片请添加！<% else %>[简介] <%= per_brief %>；共 <font color="red"><%= per_totle %></font> 张照片。
<% end if %></td>
              </tr>

              <tr  id="p_group" style="display:<% if not i=1 then %><% end if %>">
 		<td><iframe id="pho_if1" name="pho_if1" height="245" scrolling="no" frameborder="0" width="100%" src="z_photo_shares.asp?group_id=<%= per_id %>"></iframe></td>
              </tr>

            </table>
<br><br>
 <% 

loop
end if
rs.close
%> 
            <br>
            </td>
        </tr>
      </table>


      <table width="100%" border="0" align=center cellpadding="0" cellspacing="0" class=tableborder1>
        <tr> 
          <td height="22" align=right valign=middle class=TopLighNav1> 
            <div align="center"><font color=red>个人相册区</font></div></td>
        </tr>
        <tr > 
          <td  class=tablebody2><br> 
            <% 
dim j,per_shared
set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo_group where pho_all='n' and pho_state='v' and username='"&membername&"' order by pho_shared desc"
    rs.open sql,z_photo_cn,1,3
if rs.eof then

    Response.Write "<font color=red> 您还没有建立一个个人相册，请到相册组管理页面中创立！</font>"

else
do while not rs.eof 
per_id=rs("pho_group_id")
per_name=rs("pho_group_name")
per_brief=rs("pho_group_brief")
per_totle=rs("pho_group_totle")
per_shared=rs("pho_shared")
rs.movenext
i=i+1
j=j+1
 %>
            <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder1>
              <tr onClick="show(<%= i %>)" style="cursor:hand"> 
                <th height="25" width="97%"><div align="left">-=&gt; <% if (per_shared="y") then%><font color="red">共享：</font><% end if%><%= per_name %> 
                    <font id=f_flag face="Webdings">6</font></div></th>
              </tr>
              <tr> 
                <td height="20" class=tablebody2>
                  <% if per_totle<1 then %>
                  该相册中还没有图片请添加！
                  <% else %>
                  [简介] <%= per_brief %>；共 <font color="red"><%= per_totle %></font> 
                  张照片。 
                  <% end if %>
                </td>
              </tr>
              <tr  id="p_group" style="display:<% if not j=1 then %><% end if %>"> 
                <td><iframe id="pho_if" height="245" scrolling="no" frameborder="0" width="100%" src="z_photo_shares.asp?group_id=<%= per_id %>"></iframe></td>
              </tr>
            </table>
            <br>
            <br> 
            <% 

loop
end if
rs.close
%>
            <br> </td>
        </tr>
      </table><p></p></tr>
            </table></td>
        </tr>
      </table>

	</td></tr></table>

<% 

	end sub%>