<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
stats="��������"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�в鿴���������������Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()%>
<TABLE bgColor=#fff7f7 border=0 cellPadding=0 cellSpacing=0 width=760 align="center" height="402">
  <TBODY> 
  <TR vAlign=top>
    <TD width="40" height="287" background="images/img_marry/hf.jpg">       
 <script language="">
//����marquee�Ŀ��(in pixels)
var marqueewidth=360
//����marquee�ĸ߶�
var marqueeheight=350
//����marquee���ٶ�
var speed=1
//����marquee����ʾ����
/////////////////////////////////////////////////////////////////


var marqueecontents='<left><SPAN STYLE="FONT-FAMILY: ����; FONT-SIZE: 9pt; LINE-HEIGHT:15px"><B><center>��һ�� �� ��</center></b><BR>��һ�� ������<%=forum_info(0)%>������ͥ��ϵ�Ļ���׼��.<br>�ڶ��� ������������������ϵֻ����������Ч,��Ů˫��������<br>&nbsp;&nbsp;&nbsp;&nbsp;���������� ���������������ʵ�������޹ء���������δ�ر�������������ؾ�ָ�������ڣ�<BR>������  ʵ�л������ɡ�һ��һ�ޡ���Ůƽ�ȵĻ����ƶȡ�<BR>������ ��ֹ���졢������������������������ɵ���Ϊ����ֹ���<br>&nbsp;&nbsp;&nbsp;&nbsp;����ȡ����,Υ��û ����ȫ���Ʋ���������һ�����ϡ���ֹ��<br>&nbsp;&nbsp;&nbsp;&nbsp;�顣��ֹͬһ��ע�������������û������Լ����Լ���顣  <BR>������ ����Ӧ��������ʵ���������أ���ͥ��Ա��Ӧ�����ϰ��ף�<br>&nbsp;&nbsp;&nbsp;&nbsp;���������ά��ƽ�ȡ������������Ļ�����ͥ��ϵ�� <BR><br><B><center>�ڶ��� �� ��</center></B><BR>������  ��������Ů˫����ȫ��Ը�������κ�һ������������ǿ��<br>&nbsp;&nbsp;&nbsp;&nbsp;���κε����߼��� ���档<BR>������ ����Ů�����������������ߣ����������飺<BR>&nbsp;&nbsp;&nbsp;&nbsp;(һ)˫������������ע ����� <BR>&nbsp;&nbsp;&nbsp;&nbsp;(��)˫�����������Ա�ͬ��<BR>&nbsp;&nbsp;&nbsp;&nbsp;(��)˫��Ŀǰ���ǵ�������δ��ĺ��Ѿ�����ģ���<BR>&nbsp;&nbsp;&nbsp;&nbsp;(��)���˫����������һ�����������ʻ����г�������������<br>&nbsp;&nbsp;&nbsp;&nbsp;֧�� ����ķ��á�<BR>�ڰ��� ����������֮һ�ģ�������Ч��<BR>&nbsp;&nbsp;&nbsp;&nbsp;(һ)�ػ�ģ� <BR>&nbsp;&nbsp;&nbsp;&nbsp;(��)ͬ�Ա�ģ� <BR>&nbsp;&nbsp;&nbsp;&nbsp;(��)�Լ����Լ��Ĵ��׽�顣 <BR>�ھ���  Ҫ�������Ů˫���������Ե������Ļ����ǼǴ�����<br>&nbsp;&nbsp;&nbsp;&nbsp;���Ǽǣ������� һ���Ľ�������С����ϱ����涨������<br>&nbsp;&nbsp;&nbsp;&nbsp;�Ǽǣ��������֤�� <BR>��ʮ��  �Ǽǽ���ȡ�ý��֤����ȷ�����޹�ϵ����ʱ����˫��<br>&nbsp;&nbsp;&nbsp;&nbsp;�ļ��н�������һ ��ͨ�������˫��������Ϣ����ʾ����״��Ϊ���ѻ顱��ͬʱ��ʾ��ż�ǳơ� <BR><B><center>������ ��ͥ��ϵ</center></B><BR>��ʮһ�� �����ڼ�ͥ�е�λƽ�ȡ�<BR>��ʮ����  ����˫�����и����Լ�������Ȩ����<BR>��ʮ���� ����˫������ʵ�мƻ�����������<BR>��ʮ����  ����˫���ĲƲ���Ŀǰ�ڻ���Թ�������У���˫�����Դ���һ������Ϊ�� ���е�ծ��<BR>��ʮ���� �����л�����ֵ�����<BR><br><B><center>������ �� ��</center></B><BR>&nbsp;&nbsp;�ڶ�ʮ���� Υ�������ߣ���ʵ���������������Ӧ���������ֻ����Ʋá� <BR>&nbsp;&nbsp;�ڶ�ʮ���� �����в����Ƶĵط����᲻���������޶����������������й涨��<BR> &nbsp;&nbsp;�ڶ�ʮ���� ������2002��10��23����ʩ�С�<A href="#">��ҳ��</A></span></center>'
//////////////////////////////////////////////////////////////////
if (document.all)
document.write('<marquee direction="up" scrollAmount='+speed+' style="width:'+marqueewidth+';height:'+marqueeheight+'">'+marqueecontents+'</marquee>')

function regenerate(){
window.location.reload()
}
function regenerate2(){
if (document.layers){
setTimeout("window.onresize=regenerate",450)
intializemarquee()
}
}

function intializemarquee(){
document.cmarquee01.document.cmarquee02.document.write(marqueecontents)
document.cmarquee01.document.cmarquee02.document.close()
thelength=document.cmarquee01.document.cmarquee02.document.height
scrollit()
}

function scrollit(){
if (document.cmarquee01.document.cmarquee02.top>=thelength*(-1)){
document.cmarquee01.document.cmarquee02.top-=speed
setTimeout("scrollit()",100)
}
else{
document.cmarquee01.document.cmarquee02.top=marqueeheight
scrollit()
}
}

window.onload=regenerate2
</script>
       </TD></TR>              
  </TBODY></TABLE>                  
<%end sub%>