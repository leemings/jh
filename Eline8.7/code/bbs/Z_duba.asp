<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%


	stats="��������"
	call nav()
	call head_var(2,0,"","")

	dim msgtext,savestats
	
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳ʹ�ö������ߵ�Ȩ�ޣ����ȵ�½����ͬ����Ա��ϵ��"
		call dvbbs_error()	
	else
		call main()
	end if
	call activeonline()
	call footer()

sub main()



%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TR><Th  colSpan=2 align=left height=25>-=>�ʰ�����</Th></TR>
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
      <p align="center">��ɽ�������߲鶾Ϊ��ɽ��˾��Ȩ���У��κ��˶�����ʹ�����߲鶾���񣬶����븶�ѡ�<br>
    <li>
      <p align="center">������ǵ�һ��ʹ�����߲鶾���񣬻��������ϴη��ʺ����߲鶾������������ô������Ҫ���ĵȴ�����ļ����ص������ؼ�����ϡ�<br>
    <li>
      <p align="center">һ������£�ÿ��һ�ν�ɽ���Ի����������⣬�������²������������û�������У���ʱ�����߲鶾�Ĳ�����Ҳ������Ӧ������<br>
    <li>
      <p align="center">���߲鶾���û�����Ҫ�ֹ�����������ֻҪ�������߲鶾ҳ�棬�ͻ������°汾�����߲鶾��<br>
    <li>
      <p align="center">ȱʡ����£����߲鶾ɨ�豾������Ӳ�̣��������Ҫָ��ɨ���·�������ֶ�ѡ��<br>
    <li>
      <p align="center">���&quot;�������߲鶾&quot;��ť���������߲鶾ҳ�棬�����ļ������һ��ʱ�䣬�����ĵȺ�<br>
    <li>
      <p align="center">���&quot;�������߲鶾&quot;��ť��û��������ʾ�������ָ�ҳ�޷���ʾ����������һ�������½��롣<br>
    <li>
      <p align="center">������Ȩ���κ���֯����˲��ý���ɽ�������߲鶾���������á�<br>
    <li>
      <p align="center">��ɽ�������߲鶾,����Ϊ��ʽ�汾,���������ɽ����ͬʱ���������߲鶾Ŀǰֻ�鶾����ɱ�������������ɱ�����뵽����.<br>
    <li>
      <p align="center">������Խ�ɽ�������߲鶾��ʲô����ͽ���,������ϵ  
      <a href="mailto:chenxiuzhang@kingsoft.net">chenxiuzhang@kingsoft.net</a><br>
    <li>
      <p align="center">��ɽ�������߲鶾Ϊ��ɽ���Ե�һ��ս�Է�չ����,��ɽ���Խ�Ͷ��޴�������������з�,������IE�¾��ܲ�ɱ���ֲ�����ľ���������.�����������ϵ��ɽ�������߲鶾ҵ��,��<a href="http://www.duba.net/support/contact/index.htm">��ϵ����</a>��!лл������<br>
    </li>
  </ol>
  <div align="center">
    <p align="center">
    <a onmouseover="MM_swapImage('Image2','','http://www.duba.net/image/v4/a/Go_Btn_o.gif',1)" onmouseout="MM_swapImgRestore()" href="http://www.duba.net/antiscan/scan.htm"><br>
    <br>
    </a><a onmouseover="MM_swapImage('Image3','','http://www.duba.net/image/v4/a/Go_Btn_o.gif',1)" onmouseout="MM_swapImgRestore()" href="http://www.duba.net/antiscan/scan.htm"><img src="http://www.duba.net/image/v4/a/Go_Btn.gif" border="0" name="Image3" width="123" height="34"></a>
  </div>
  <p align="center">��
  </td>
      </table>
<%
end sub
%>