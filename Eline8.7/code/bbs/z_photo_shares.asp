<!--#include file="conn.asp"-->
<!--#include file="z_photo_conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/Forum_css.asp"-->
<script>
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script>
<script language="JavaScript">
nereidFadeObjects = new Object();
nereidFadeTimers = new Object();
function nereidFade(object, destOp, rate, delta){
if (!document.all)
return
    if (object != "[object]"){  //do this so I can take a string too
        setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
        return;
    }
    clearTimeout(nereidFadeTimers[object.sourceIndex]);
    diff = destOp-object.filters.alpha.opacity;
    direction = 1;
    if (object.filters.alpha.opacity > destOp){
        direction = -1;
    }
    delta=Math.min(direction*diff,delta);
    object.filters.alpha.opacity+=direction*delta;
    if (object.filters.alpha.opacity != destOp){
        nereidFadeObjects[object.sourceIndex]=object;
        nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
    }
}

</script>
<script language="JavaScript" type="text/JavaScript">
function picsize(n)
{
var a,b,c,d
a=document.all["z_photo_pic"][n].width;
b=document.all["z_photo_pic"][n].height;
c=a/b;
d=160/95;
if (c>d)
{
if (document.all["z_photo_pic"][n].width>160)
document.all["z_photo_pic"][n].width=160;
}
else {
if (document.all["z_photo_pic"][n].height>95)
document.all["z_photo_pic"][n].height=95;
}
}
</script>

<script language="JavaScript" type="text/JavaScript">
function picsize1()
{
var a,b,c,d
a=document.all["z_photo_pic1"].width;
b=document.all["z_photo_pic1"].height;
c=a/b;
d=160/95;
if (c>d)
{
if (document.all["z_photo_pic1"].width>160)
document.all["z_photo_pic1"].width=160;
}
else {
if (document.all["z_photo_pic1"].height>95)
document.all["z_photo_pic1"].height=95;
}
}
</script>
<script language="JavaScript" type="text/JavaScript">
function picsize2()
{
var a,b,c,d
a=document.all["z_photo_pic2"].width;
b=document.all["z_photo_pic2"].height;
c=a/b;
d=160/95;
if (c>d)
{
if (document.all["z_photo_pic2"].width>160)
document.all["z_photo_pic2"].width=160;
}
else {
if (document.all["z_photo_pic2"].height>95)
document.all["z_photo_pic2"].height=95;
}
}
</script>
<script language="JavaScript" type="text/JavaScript">
function picsize3()
{
var a,b,c,d
a=document.all["z_photo_pic3"].width;
b=document.all["z_photo_pic3"].height;
c=a/b;
d=160/95;
if (c>d)
{
if (document.all["z_photo_pic3"].width>160)
document.all["z_photo_pic3"].width=160;
}
else {
if (document.all["z_photo_pic3"].height>95)
document.all["z_photo_pic3"].height=95;
}
}
</script>
<script language="JavaScript" type="text/JavaScript">
function picsize4()
{
var a,b,c,d
a=document.all["z_photo_pic4"].width;
b=document.all["z_photo_pic4"].height;
c=a/b;
d=160/95;
if (c>d)
{
if (document.all["z_photo_pic4"].width>160)
document.all["z_photo_pic4"].width=160;
}
else {
if (document.all["z_photo_pic4"].height>95)
document.all["z_photo_pic4"].height=95;
}
}
</script>


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="picsize1();picsize2();picsize3();picsize4()">
<%
dim p_id

p_id=request.querystring("p_id")
if p_id="" then 
else 
'	set rs=server.createobject("adodb.recordset")
'	sql="select pho_count from z_photo where pho_id = "&p_id
'	rs.open sql,z_photo_cn,1,3
'	rs("pho_count")=rs("pho_count")+1
'	rs.update
'	rs.close
end if

stats="个人相册"

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入个人相册的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else
	call start()
end if
if founderr then  call dvbbs_error()

'=========================================================
sub start()
  

'-------------------------
%>

  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
<table width="100%" border="0" align=center cellpadding="0" cellspacing=0 >
        <% 
dim group_id,pho_group_name
group_id=request.QueryString("group_id")
group_id=cint(group_id)
 
 %>
        <% 
dim pho_url1,pho_title1,pho_brief1,pho_adtime1,pho_id1,username1,pho_count1,p1
dim pho_url2,pho_title2,pho_brief2,pho_adtime2,pho_id2,username2,pho_count2,p2
dim pho_url3,pho_title3,pho_brief3,pho_adtime3,pho_id3,username3,pho_count3,p3
dim pho_url4,pho_title4,pho_brief4,pho_adtime4,pho_id4,username4,pho_count4,p4
dim ps,mx,mp,cp,page
ps=4

page=request.querystring("page")
if page="" then
cp=1
else
cp=cint(page)
end if

 set rs=server.createobject("adodb.recordset")
    sql="select * from z_photo where (state = 'v' and pho_group ='"&group_id&"' ) order by pho_id desc "
    rs.open sql,z_photo_cn,1,3
if rs.eof then
	Response.Write "<tr>" 
    Response.Write "<td height=210 valign=top class=tablebody2>目前还没有照片！<img id=z_photo_pic1 style=display:none></img></td>"
	response.Write"<img id=z_photo_pic2 style=display:none></img><img id=z_photo_pic3 style=display:none></img><img id=z_photo_pic4 style=display:none></img>"
    Response.Write "</tr>"
else
mx=rs.recordcount
mp=cint(mx mod ps)
if mp=0 then
mp=cint(mx\ps)
else

mp=cint(mx\ps)+1
end if

rs.Pagesize=ps
if mp<cp then
rs.absolutepage=mp
else
rs.absolutepage=cp
end if
if not rs.eof then

pho_id1=rs("pho_id")
username1=rs("username")
pho_url1=rs("pho_url")
pho_adtime1=rs("pho_adtime")
pho_adtime1=formatdatetime(pho_adtime1,2)
pho_title1=rs("pho_title")
if (lenb(pho_title1)>20) then
pho_title1=left(pho_title1,20)&"..."
end if
pho_brief1=rs("pho_brief")
pho_count1=rs("pho_count")

end if
rs.movenext
if rs.eof then
rs.moveprevious
p2="none"
else
pho_id2=rs("pho_id")
username2=rs("username")
pho_url2=rs("pho_url")
pho_adtime2=rs("pho_adtime")
pho_adtime2=formatdatetime(pho_adtime2,2)
pho_title2=rs("pho_title")
if (lenb(pho_title2)>20) then
pho_title2=left(pho_title2,20)&"..."
end if
pho_brief2=rs("pho_brief")
pho_count2=rs("pho_count")
end if

rs.movenext
if rs.eof then
rs.moveprevious
p3="none"
else
pho_id3=rs("pho_id")
username3=rs("username")
pho_url3=rs("pho_url")
pho_adtime3=rs("pho_adtime")
pho_adtime3=formatdatetime(pho_adtime3,2)
pho_title3=rs("pho_title")
if (lenb(pho_title3)>20) then
pho_title3=left(pho_title3,20)&"..."
end if
pho_brief3=rs("pho_brief")
pho_count3=rs("pho_count")
end if

rs.movenext
if rs.eof then
rs.moveprevious
p4="none"
else
pho_id4=rs("pho_id")
username4=rs("username")
pho_url4=rs("pho_url")
pho_adtime4=rs("pho_adtime")
pho_adtime4=formatdatetime(pho_adtime4,2)
pho_title4=rs("pho_title")
if (lenb(pho_title4)>20) then
pho_title4=left(pho_title4,20)&"..."
end if
pho_brief4=rs("pho_brief")
pho_count4=rs("pho_count")
end if 


 %>
        <tr> 
          <td class=tablebody2><br> <table width="660" align="center" cellpadding="1" cellspacing="3"  >
              <tr> 
                <td><div align="left"><strong><%= pho_title1 %></strong></div></td>
                <td style="display:<%= p2 %>"><div align="left"><strong><%= pho_title2 %></strong></div></td>
                <td style="display:<%= p3 %>"><div align="left"><strong><%= pho_title3 %></strong></div></td>
                <td style="display:<%= p4 %>"><div align="left"><strong><%= pho_title4 %></strong></div></td>
              </tr>
              <tr> 
                <td height="95" class=tablebody2><div align="left">
                    <a href="javascript:openScript('z_photo_view1.asp?pho_id=<%=pho_id1%>&action=new',550,590)" onClick="window.location='z_photo_shares.asp?group_id=<%= group_id %>&page=<%=cp%>&p_id=<%=pho_id1%>'"><img src="<%= pho_url1%>" border="0" id="z_photo_pic1" onload="picsize1()" onMouseOut=nereidFade(this,45,50,5) onMouseOver=nereidFade(this,100,0,5)  style="FILTER: alpha(opacity=45)" ></a>
				</div>
				</td>
     			 <td height="95" class=tablebody2 style="display:<%= p2 %>"><div align="left">
                     <a href="javascript:openScript('z_photo_view1.asp?pho_id=<%=pho_id2%>&action=new',550,590)" onClick="window.location='z_photo_shares.asp?group_id=<%= group_id %>&page=<%=cp%>&p_id=<%=pho_id2%>'"><img id="z_photo_pic2" src="<%= pho_url2%>" border="0" onload="picsize2()" onMouseOut=nereidFade(this,45,50,5) onMouseOver=nereidFade(this,100,0,5)  style="FILTER: alpha(opacity=45)" ></a>
				</div>
				</td>
                     
    			  <td height="95" class=tablebody2 style="display:<%= p3 %>"><div align="left">
                    <a href="javascript:openScript('z_photo_view1.asp?pho_id=<%=pho_id3%>&action=new',550,590)" onClick="window.location='z_photo_shares.asp?group_id=<%= group_id %>&page=<%=cp%>&p_id=<%=pho_id3%>'"><img id="z_photo_pic3" src="<%= pho_url3%>" alt=""  border="0" onload="picsize3()" onMouseOut=nereidFade(this,45,50,5) onMouseOver=nereidFade(this,100,0,5)  style="FILTER: alpha(opacity=45)" ></a>
                  </div></td>
    			  <td height="95" class=tablebody2 style="display:<%= p4 %>"><div align="left">
                   <a href="javascript:openScript('z_photo_view1.asp?pho_id=<%=pho_id4%>&action=new',550,590)" onClick="window.location='z_photo_shares.asp?group_id=<%= group_id %>&page=<%=cp%>&p_id=<%=pho_id4%>'"><img id="z_photo_pic4" src="<%= pho_url4%>" alt="" border="0" onload="picsize4()" onMouseOut=nereidFade(this,45,50,5) onMouseOver=nereidFade(this,100,0,5)  style="FILTER: alpha(opacity=45)" ></a>
              </tr>
              <tr> 
                <td><div >作者：<%= username1 %></div></td>
                <td  style="display:<%= p2 %>"><div >作者：<%= username2 %></div></td>
                <td  style="display:<%= p3 %>"><div >作者：<%= username3 %></div></td>
                <td  style="display:<%= p4 %>"><div >作者：<%= username4 %></div></td>
              </tr>
				<tr> 
                <td><div >时间：<%= pho_adtime1 %></div></td>
                <td style="display:<%= p2 %>"><div >时间：<%= pho_adtime2 %></div></td>
                <td style="display:<%= p3 %>"><div >时间：<%= pho_adtime3 %></div></td>
                <td style="display:<%= p4 %>"><div >时间：<%= pho_adtime4 %></div></td>
              </tr>
				<tr> 
                <td ><div >点击数：<font color="#FF0000"><%= pho_count1 %></font></div></td>
			    <td style="display:<%= p2 %>" ><div >点击数：<font color="#FF0000"><%= pho_count2 %></font></div></td>
			    <td style="display:<%= p3 %>" ><div >点击数：<font color="#FF0000"><%= pho_count3 %></font></div></td>
			    <td style="display:<%= p4 %>" ><div >点击数：<font color="#FF0000"><%= pho_count4 %></font></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align=center class=tablebody2><BR>共<font color="red"> <%= mx %> </font>张，<font color="red"><%= mp %> </font>页 : <%call DispPageNum(cp,mp,"'?group_id="&group_id&"&page=","'")%><BR><BR></td>
        </tr>
        <%
 end if %>
      </table>
	</td>
  </tr>
</table>


<% 
call close_z_photo_cn()
	end sub%>
