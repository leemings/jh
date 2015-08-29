<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
stats="婚姻法规"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有查看本社区婚姻法规的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
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
//设置marquee的宽度(in pixels)
var marqueewidth=360
//设置marquee的高度
var marqueeheight=350
//设置marquee的速度
var speed=1
//设置marquee的显示内容
/////////////////////////////////////////////////////////////////


var marqueecontents='<left><SPAN STYLE="FONT-FAMILY: 宋体; FONT-SIZE: 9pt; LINE-HEIGHT:15px"><B><center>第一章 总 则</center></b><BR>第一条 本法是<%=forum_info(0)%>婚姻家庭关系的基本准则.<br>第二条 本婚姻法及本婚姻关系只在社区内有效,男女双方所建立<br>&nbsp;&nbsp;&nbsp;&nbsp;的是社区内 的虚拟婚姻，与现实社会婚姻无关。（以下若未特别声明，婚姻相关均指本社区内）<BR>第三条  实行婚姻自由、一夫一妻、男女平等的婚姻制度。<BR>第四条 禁止包办、买卖婚姻和其他干涉婚姻自由的行为。禁止借婚<br>&nbsp;&nbsp;&nbsp;&nbsp;姻索取财物,违者没 收其全部财产，并入狱一周以上。禁止重<br>&nbsp;&nbsp;&nbsp;&nbsp;婚。禁止同一人注册两个社区的用户名，自己和自己结婚。  <BR>第五条 夫妻应当互相忠实，互相尊重；家庭成员间应当敬老爱幼，<br>&nbsp;&nbsp;&nbsp;&nbsp;互相帮助，维护平等、和睦、文明的婚姻家庭关系。 <BR><br><B><center>第二章 结 婚</center></B><BR>第六条  结婚必须男女双方完全自愿，不许任何一方对他方加以强迫<br>&nbsp;&nbsp;&nbsp;&nbsp;或任何第三者加以 干涉。<BR>第七条 当男女两方符合以下条件者，方可申请结婚：<BR>&nbsp;&nbsp;&nbsp;&nbsp;(一)双方均是社区的注 册居民； <BR>&nbsp;&nbsp;&nbsp;&nbsp;(二)双方在社区的性别不同；<BR>&nbsp;&nbsp;&nbsp;&nbsp;(三)双方目前均是单身（包括未婚的和已经离异的）；<BR>&nbsp;&nbsp;&nbsp;&nbsp;(四)结婚双方（或其中一方）的社区帐户里有充足的人民币足以<br>&nbsp;&nbsp;&nbsp;&nbsp;支付 其结婚的费用。<BR>第八条 有下列情形之一的，婚姻无效：<BR>&nbsp;&nbsp;&nbsp;&nbsp;(一)重婚的； <BR>&nbsp;&nbsp;&nbsp;&nbsp;(二)同性别的； <BR>&nbsp;&nbsp;&nbsp;&nbsp;(三)自己与自己的大米结婚。 <BR>第九条  要求结婚的男女双方必须亲自到社区的婚姻登记处进行<br>&nbsp;&nbsp;&nbsp;&nbsp;结婚登记，并依照 一定的结婚程序进行。符合本法规定的予以<br>&nbsp;&nbsp;&nbsp;&nbsp;登记，发给结婚证。 <BR>第十条  登记结婚后，取得结婚证，即确定夫妻关系。届时夫妻双方<br>&nbsp;&nbsp;&nbsp;&nbsp;的家中将会增开一 条通道。结婚双方个人信息将显示婚姻状况为“已婚”，同时显示配偶昵称。 <BR><B><center>第三章 家庭关系</center></B><BR>第十一条 夫妻在家庭中地位平等。<BR>第十二条  夫妻双方都有各用自己姓名的权利。<BR>第十四条 夫妻双方都有实行计划生育的义务。<BR>第十五条  夫妻双方的财产，目前在婚后仍归各自所有，由双方各自处理，一方无需为对 方承担债务。<BR>第十六条 夫妻有互相扶持的义务。<BR><br><B><center>第四章 附 则</center></B><BR>&nbsp;&nbsp;第二十三条 违反本法者，依实际情况，将予以相应的行政处分或法律制裁。 <BR>&nbsp;&nbsp;第二十四条 本法中不完善的地方将会不断作补充修订或由社区政府另行规定。<BR> &nbsp;&nbsp;第二十五条 本法自2002年10月23日起施行。<A href="#">回页首</A></span></center>'
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