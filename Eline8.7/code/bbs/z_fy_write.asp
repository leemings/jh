<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_write.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
Response.Buffer=True 
stats="撰写状纸"
call nav()
call head_var(0,0,"社区法院","z_fy_fayuan.asp")
session("username")=membername
if not founduser then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区法院的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
call dvbbs_error()
else
  response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>撰写状子</b></td></tr><tr><td valign=middle class=tablebody2 height=100 ><CENTER>"
  call getfyconfig()
  response.write "<form action=z_fy_newplan.asp method=post  width=60% class=tablebody1>"
 response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center >"
 response.write "<tr><td class=tablebody2 rowspan=3 width=10% > 注意事项：<br><br>①投诉主题要求简明扼要，最好限制在20个汉字内。<br>②对被告的处罚要求请根据社区条例确定，并注意尺度。<br>③投诉内容和理由请详细填写相关证据，如果涉及论坛发帖违例情况，请标明帖子位置。</td>"
 response.write "<td colspan=2 width=20% class=tablebody1><b>投诉主题：</b><input name=topic size=86 maxlength=86></td></tr>"
 response.write "<tr><td width=15%  class=tablebody1><b>被　　告：</b><input name=towho size=22 maxlength=22></td>"
 response.write "<td width=20%  class=tablebody1><b>要求对被告进行：</b>" 
 response.write "<select name=play size=1>"
 response.write "<option value='终身服刑'><b>╋≡终身服刑≡</b></option>"
 response.write "<option value='入狱10分钟'><b>&nbsp;&nbsp;├入狱10分钟</b></option>"
 response.write "<option value='入狱半小时'><b>&nbsp;&nbsp;├入狱半小时</b></option>"
 response.write "<option value='入狱一小时'><b>&nbsp;&nbsp;├入狱一小时</b></option>"
 response.write "<option value='入狱一天'><b>&nbsp;&nbsp;├入狱一天</b></option>"
 response.write "<option value='入狱三天'><b>&nbsp;&nbsp;├入狱三天</b></option>"
 response.write "<option value='入狱一周'><b>&nbsp;&nbsp;├入狱一周</b></option>"
 response.write "<option value='没收全部财产'><b>╋≡没收全部财产≡</b></option>" 
 response.write "<option value='罚款1％'><b>&nbsp;&nbsp;├罚款1%</b></option>"
 response.write "<option value='罚款10％'><b>&nbsp;&nbsp;├罚款10%</b></option>"
 response.write "<option value='罚款50％'><b>&nbsp;&nbsp;├罚款50%</b></option>"
 response.write "<option value='扣除所有经验'><b>╋≡扣除所有经验≡</b></option>"               
 response.write "<option value='扣经验1％'><b>&nbsp;&nbsp;├扣经验1%</b></option>"
 response.write "<option value='扣经验10％'><b>&nbsp;&nbsp;├扣经验10%</b></option>"               
 response.write "<option value='扣经验50％'><b>&nbsp;&nbsp;├扣经验50%</b></option>"
 response.write "<option value='扣除所有魅力'><b>╋≡扣除所有魅力≡</b></option>" 
 response.write "<option value='扣魅力1％'><b>&nbsp;&nbsp;├扣魅力1%</b></option>"               
 response.write "<option value='扣魅力10％'><b>&nbsp;&nbsp;├扣魅力10%</b></option>"               
 response.write "<option value='扣魅力50％'><b>&nbsp;&nbsp;├扣魅力50%</b></option>"
 response.write "<option value='清空'><b>╋≡所有分值清零≡</b></option>"
 response.write "<option value='威望-10'><b>&nbsp;&nbsp;├威望-10</b></option>"
 response.write "<option value='威望-50'><b>&nbsp;&nbsp;├威望-50</b></option>"
 response.write "<option value='威望-100'><b>&nbsp;&nbsp;├威望-100</b></option>"
 response.write "<option value='午门斩首'><b>╋≡午门斩首≡</b></option>"
 if lh_set(0)="1" then response.write "<option value='离婚'><b>╋≡与被告离婚≡</b></option>"
 response.write "</select> 的处罚</td></tr>"
 response.write "<tr><td colspan=2 align=left class=tablebody1><b>投诉内容和理由:</b><br><textarea name=text cols=100 rows=8></textarea><br></td></tr>"
 response.write "<tr><td align=center colspan=3 class=tablebody2>"
 response.write "<input type=submit value='告　他' name=submit >&nbsp;&nbsp;&nbsp;&nbsp;" 
 response.write "<input type=reset value='重　写' name=reset >&nbsp;&nbsp;&nbsp;&nbsp;"
 response.write "<input type=button value='返　回' onClick='javascript:history.back()' name=button>"
 response.write "<br></td></tr></table></form></CENTER></td></tr></table>"     
end if  
call footer()  
%> 