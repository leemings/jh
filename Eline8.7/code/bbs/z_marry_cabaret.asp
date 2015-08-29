<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<!--#include file="inc/FORUM_CSS.asp"-->
<%
if request("name")<>"" and request("id")<>"" then
	name=request("name")
	id=request("id") 
else
	name="一生一意"
	id="1"
end if

stats="社区大酒店"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区大酒店的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call dvbbs_error()
else
	call main()
end if
call activeonline()
call footer()

sub main()
%>
<!--<HTML><HEAD><TITLE>社区大酒店</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type> -->
<SCRIPT language=JavaScript>
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

</SCRIPT>
<!--<link rel="stylesheet" href="Forum.css" type="text/css">
</HEAD>-->
<BODY bgColor=#ffffff leftMargin=0 
onload="MM_preloadImages('images/img_marry/index_1h.gif','images/img_marry/index_2h.gif','images/img_marry/index_3h.gif','images/img_marry/index_4h.gif','images/img_marry/index_5h.gif','images/img_marry/index_6h.gif','images/img_marry/index_7h.gif','images/img_marry/index_8h.gif','images/img_marry/index_9h.gif')" 
text=#000000 topMargin=0>
<%
call marry_nav()
%><TABLE border=0 cellPadding=0 cellSpacing=0 width=770 align="center" height="403">
  <TBODY> 
  <TR>
    <TD class=ct-def2 height=40 width=100><img border="0" src="images/img_marry/left.jpg"></TD>
    <TD class=ct-def2 height=40 width=346>
    <img border="0" src="images/img_marry/middle.jpg"></TD>
    <TD bgColor=#F82828 class=ct-def2 height=40 width=324>
    <img border="0" src="images/img_marry/right.jpg"></TD>
  </TR>
  <TR>
    <TD background=images/img_marry/jiubj.gif colSpan=3 
    height=25 width="770">　</TD></TR>
  <TR>
    <TD colSpan=3 height="158" width="770">
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=769 height="157">
        <TBODY>
        <TR>
          <TD height=157 width=123 valign="top">
          <img border="0" src="images/img_marry/girl.gif"></TD>
          <TD height=157 width=550 background="images/img_marry/jhbj.gif">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="146">
            <TR>
                      <TD height=117><A 
                        href="?name=一心一意&id=1" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('a','','images/img_marry/index_1h.gif',1)" 
                        target=list><IMG border=0 height=208 name=a 
                        src="images/img_marry/index_1.gif" width=38></a></TD>
                      <TD><A 
                        href="?name=比翼双飞&id=2" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('b','','images/img_marry/index_2h.gif',1)" 
                        target=list><IMG border=0 height=208 name=b 
                        src="images/img_marry/index_2.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=三生有幸&id=3" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('c','','images/img_marry/index_3h.gif',1)" 
                        target=list><IMG border=0 height=208 name=c 
                        src="images/img_marry/index_3.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=海誓山盟&id=4" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('d','','images/img_marry/index_4h.gif',1)" 
                        target=list><IMG border=0 height=208 name=d 
                        src="images/img_marry/index_4.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=五彩缤纷&id=5" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('e','','images/img_marry/index_5h.gif',1)" 
                        target=list><IMG border=0 height=208 name=e 
                        src="images/img_marry/index_5.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=流光飞舞&id=6" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('f','','images/img_marry/index_6h.gif',1)" 
                        target=list><IMG border=0 height=208 name=f 
                        src="images/img_marry/index_6.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=郎情妾意&id=7" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('g','','images/img_marry/index_7h.gif',1)" 
                        target=list><IMG border=0 height=208 name=g 
                        src="images/img_marry/index_7.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=白头到老&id=8" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('h','','images/img_marry/index_8h.gif',1)" 
                        target=list>
                      <IMG border=0 height=208 name=h 
                        src="images/img_marry/index_8.gif" width=38></A></TD>
                      <TD><A 
                        href="?name=天长地久&id=9" 
                        onmouseout=MM_swapImgRestore() 
                        onmouseover="MM_swapImage('i','','images/img_marry/index_9h.gif',1)"
						target=list><IMG border=0 height=208 name=i 
                        src="images/img_marry/index_9.gif" 
                    width=37></A></TD></TR></td>
              </table>
          </TD>
          <TD height=157 width=96 valign="top">
          <img border="0" src="images/img_marry/boy1.gif"></TD>
          </TR>
        </TBODY></TABLE></TD></TR>
  <TR>
    <TD align=right background=images/img_marry/index_bg1.gif class=ct-def2 colSpan=3 height=20 width="770">
		<FONT color=#ffffff>正在进行的婚礼酒席＊</FONT>&nbsp;&nbsp;</TD></TR>
  <TR>
    <TD bgColor=#fcc182 colSpan=3 height=16 vAlign=center width="770">
		&nbsp;&nbsp;<img src="" height=14 width=0><%=name%>
		<table width="770" border="0" cellspacing="0" cellpadding="0" class=ct-def2>
		<%  
		set rs1=server.createobject("adodb.recordset")

		rs1.open"select * from jie where type="&id&" and datediff('h',year,'"&now()&"')<=long",connj,1,1
	  	if rs1.bof then 
			rs1.close  %>
			<tr class=ct-def2> 
				<td bgcolor=D70000 width=100% align=center valign=middle><img src="" height=23 width=0>现在没有人结婚</td>
			</tr>
		<% else %>
		  <tr class=ct-def2>
		    <td bgcolor=D70000 width=45>&nbsp;</td>
		    <td bgcolor=D70000 width=150 align=left><font color=#FFFFFF>新郎</font>&nbsp;<img src=images/img_marry/index_male.gif width=25 height=23></td>
		    <td bgcolor=D70000 width=150 align=left><font color=#FFFFFF>新娘</font>&nbsp;<img src=images/img_marry/index_female.gif width=25 height=23></td>
		    <td bgcolor=D70000 width=160 align=left><font color=#FFFFFF>开始时间</font>&nbsp;<img src=images/img_marry/index_time.gif width=22 height=23></td>
		    <td bgcolor=D70000 width=160 align=left><font color=#FFFFFF>结束时间</font>&nbsp;<img src=images/img_marry/index_time.gif width=22 height=23></td>
		    <td bgcolor=D70000 width=50 align=left>&nbsp;</td> 
		    <td bgcolor=D70000 width=45>&nbsp;</td>
		  </tr>
		  <%do while not rs1.eof %>
			    <td width=45>&nbsp;</td>
			    <td align=left width=150><img src="" height=14 width=0><a href=dispuser.asp?name=<%=rs1("username")%> target="_blank" title="查看<%=rs1("username")%>的资料"><%=rs1("username")%></a></td>
				<td align=left width=150><a href=dispuser.asp?name=<%=rs1("thename")%> target="_blank" title="查看<%=rs1("thename")%>的资料"><%=rs1("thename")%></a></td>
			    <td align=left width=160><%=rs1("year")%></td>
			    <td align=left width=160><%=dateadd("h",rs1("long"),rs1("year"))%></td>
			    <td	align=left width=50><a href=?id=<%=rs1("ID")%> target=_parent title="参加婚宴">喝喜酒去</a></td>
				<td width=45 align=left>&nbsp;</td>
			  </tr>
			 <% rs1.movenext
		  loop
		  rs1.close
		  end if%>
		</table>
	</td>
  </tr>
</TBODY></table>
</body></html>
<%
end sub
%>
