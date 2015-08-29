<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%


	stats="毒霸在线"
	call nav()
	call head_var(2,0,"","")

	dim msgtext,savestats
	
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛使用毒霸在线的权限，请先登陆或者同管理员联系。"
		call dvbbs_error()	
	else
		call main()
	end if
	call activeonline()
	call footer()

sub main()



%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TR><Th  colSpan=2 align=left height=25>-=>词霸在线</Th></TR>
<TR><TD vAlign=top class=tablebody1 width=100% >
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<td width="587" valign="top" align="center" height="592">
  <div align="center">
    <p align="center">
    <a onmouseover="MM_swapImage('Image2','','http://www.duba.net/image/v4/a/Go_Btn_o.gif',1)" onmouseout="MM_swapImgRestore()" href="http://www.duba.net/antiscan/scan.htm"><img src="http://www.duba.net/image/v4/a/Go_Btn.gif" border="0" name="Image2" width="123" height="34"></a><br>
  </div>
  <ol>
    <li>
      <p align="center">金山毒霸在线查毒为金山公司版权所有，任何人都可以使用在线查毒服务，而无须付费。<br>
    <li>
      <p align="center">如果您是第一次使用在线查毒服务，或者自您上次访问后在线查毒做了升级，那么您可能要耐心等待相关文件下载到您本地计算机上。<br>
    <li>
      <p align="center">一般情况下，每周一次金山毒霸会升级病毒库，增加最新病毒的资料至用户计算机中，此时，在线查毒的病毒库也会做相应升级。<br>
    <li>
      <p align="center">在线查毒的用户不需要手工进行升级，只要来到在线查毒页面，就会是最新版本的在线查毒。<br>
    <li>
      <p align="center">缺省情况下，在线查毒扫描本地所有硬盘，如果您需要指定扫描的路径，请手动选择。<br>
    <li>
      <p align="center">点击&quot;进入在线查毒&quot;按钮将进入在线查毒页面，下载文件会持续一断时间，请耐心等候。<br>
    <li>
      <p align="center">点击&quot;进入在线查毒&quot;按钮后，没有正常显示，而出现该页无法显示，请您后退一步，重新进入。<br>
    <li>
      <p align="center">不经授权，任何组织或个人不得将金山毒霸在线查毒服务做他用。<br>
    <li>
      <p align="center">金山毒霸在线查毒,现在为正式版本,病毒库与金山毒霸同时升级，在线查毒目前只查毒，不杀毒，如果想在线杀毒，请到网易.<br>
    <li>
      <p align="center">如果您对金山毒霸在线查毒有什么意见和建议,可以联系  
      <a href="mailto:chenxiuzhang@kingsoft.net">chenxiuzhang@kingsoft.net</a><br>
    <li>
      <p align="center">金山毒霸在线查毒为金山毒霸的一个战略发展方向,金山毒霸将投入巨大的人力物力来研发,以期在IE下就能查杀各种病毒、木马、恶意程序.如果您有意联系金山毒霸在线查毒业务,请<a href="http://www.duba.net/support/contact/index.htm">联系我们</a>。!谢谢合作！<br>
    </li>
  </ol>
  <div align="center">
    <p align="center">
    <a onmouseover="MM_swapImage('Image2','','http://www.duba.net/image/v4/a/Go_Btn_o.gif',1)" onmouseout="MM_swapImgRestore()" href="http://www.duba.net/antiscan/scan.htm"><br>
    <br>
    </a><a onmouseover="MM_swapImage('Image3','','http://www.duba.net/image/v4/a/Go_Btn_o.gif',1)" onmouseout="MM_swapImgRestore()" href="http://www.duba.net/antiscan/scan.htm"><img src="http://www.duba.net/image/v4/a/Go_Btn.gif" border="0" name="Image3" width="123" height="34"></a>
  </div>
  <p align="center">　
  </td>
      </table>
<%
end sub
%>