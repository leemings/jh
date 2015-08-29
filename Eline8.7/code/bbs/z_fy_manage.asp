<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
response.buffer=true
stats="法院参数设置"
call nav()
call head_var(0,0,"社区法院","z_fy_fayuan.asp")
if not master then 
  Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区法院后台的权限，或使用了非法参数。"
   call dvbbs_error()
else
   call getfyconfig()
   if request("action")="save" then
     bs_set=request.form("bsbz")+"|"+request.form("bsmoney")
     tj_set=request.form("tjbz")+"|"+request.form("tjmoney")
     pst_set=request.form("pstbz")+"|"+request.form("pstID")
     fj_set=request.form("fj1")+","+ request.form("fj2")+","+ request.form("fj3")+","+ request.form("fj4")+","+request.form("fj5") 
     lh_set=request.form("lhbz")+"|"+request.form("lhmlbz")+"|"+request.form("lhml")
     qj_set=request.form("qjbz")+"|"+ request.form("qjmlbz")+"|"+request.form("qjml")+"|"+request.form("qjxm")+"|"+request.form("qjtime")
     log_set=request.form("logbz")+"|"+request.form("lognum")
     if isnull(request.form("fgname")) or request.form("fgname")="" then
       fyname="无"
     else
       fyname=request.form("fgname")
     end if
     if isnull(request.form("jyzname")) or request.form("jyzname")="" then
       jyname="无"
     else
       jyname=request.form("jyzname")
     end if
     sql="update [fyconfig] set fyreadme='"&request.form("text")&"',bsmoney='"&bs_set&"',tjmoney='"&tj_set&"',pst='"&pst_set&"',fj='"&fj_set&"',lh='"&lh_set&"',qj='"&qj_set&"',fyadmin='"&fyname&"',jyadmin='"&jyname&"',log='"&log_set&"'" 
     connfy.execute sql
     Response.Redirect "z_fy_fayuan.asp"
  else
   response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>法院监狱参数设置</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
   response.write "<form action=z_fy_manage.asp?action=save method=post>"
   response.write "<table class=tableborder1 cellspacing=1 cellpadding=3 align=center ><tr>"
   response.write "<th align=left colspan=2>法院功能设置</th><th align=right colspan=2>特邀法官:<input name=fgname size=18 maxlength=20 value="
   if isnull(fyname) or fyname="" then 
      response.write "无"
   else
      response.write fyname
   end if
   response.write "></th></tr>"   
   response.write "<tr><td class=tablebody2 width=15% >法院公告说明：</td><td colspan=3 class=tablebody1><textarea name=text cols=100 rows=8 >"&rs(0)&"</textarea></td>"
   response.write "<tr><td class=tablebody2>附加奖罚设定：<br>格式为 <b>原告|被告</b><br>罚款前加－号！<br>奖励前加＋号！</td><td colspan=3 class=tablebody1>同意：<input name=fj1 size=8 value="&fj_set(0)&"> 驳回：<input name=fj2 size=8 value="&fj_set(1)&"> 平反：<input name=fj3 size=8 value="&fj_set(2)&"> 诬告：<input name=fj4 size=8 value="&fj_set(3)&"> 严重诬告：<input name=fj5 size=8 value="&fj_set(4)&"></td></tr>"
   response.write "<tr><td class=tablebody2>开放陪审团功能：</td><td class=tablebody1><input type=radio name=pstbz value=1 "
   if pst_set(0)="1" then response.write " checked"
   response.write ">开放<input type=radio name=pstbz value=0 "
   if pst_set(0)="0" then response.write " checked"
   response.write ">关闭</td>"
   response.write "<td colspan=2 class=tablebody1>陪审版面BoardID：&nbsp;<input name=pstID value="&pst_set(1)&" size=18> 设定提交陪审团投票帖子发表版面ID</td></tr>"
   response.write "<tr><td class=tablebody2>开放离婚功能：</td><td class=tablebody1><input type=radio name=lhbz value=1 "
   if lh_set(0)="1" then response.write " checked"
   response.write ">开放<input type=radio name=lhbz value=0 "
   if lh_set(0)="0" then response.write " checked"
   response.write ">关闭</td>"
   response.write "<td width=15% class=tablebody1>离婚是否扣魅力？</td>"
   response.write "<td class=tablebody1><input type=radio name=lhmlbz value=1 "
   if lh_set(1)="1" then response.write " checked"
   response.write ">扣<input type=radio name=lhmlbz value=0 "
   if lh_set(1)="0" then response.write " checked"
   response.write ">不扣　  扣除魅力值<input name=lhml size=18 maxlength=20 value="&lh_set(2)&"> 是否可起诉离婚</td></tr>"   
   response.write "<th align=left colspan=2>监狱功能设置</th><th align=right colspan=2>特邀监狱长:<input name=jyzname size=18 maxlength=20 value="
   if isnull(jyname) or jyname="" then 
      response.write "无"
   else
      response.write jyname
   end if
   response.write "></th></tr>"   
   response.write "<tr><td class=tablebody2>开放保释功能：</td><td class=tablebody1><input type=radio name=bsbz value=1" 
   if bs_set(0)="1" then response.write " checked"
   response.write ">开放<input type=radio name=bsbz value=0"
   if bs_set(0)="0" then response.write " checked"
   response.write " >关闭</td>"
   response.write "<td class=tablebody1 colspan=2>最低保释金：<input name=bsmoney size=18 maxlength=20 value="&bs_set(1)&"> 设定在押囚犯获保释所需最低保释金，最高9亿</td></tr>"
   response.write "<tr><td class=tablebody2>开放探监功能：</td><td width=15% class=tablebody1><input type=radio name=tjbz value=1" 
   if tj_set(0)="1" then response.write " checked"
   response.write ">开放<input type=radio name=tjbz value=0"
   if tj_set(0)="0" then response.write " checked"
   response.write " >关闭</td>"
   response.write "<td class=tablebody1 colspan=2>探监留言费：<input name=tjmoney size=18 maxlength=20 value="&tj_set(1)&"> 设定探视在押囚犯并留言所收取现金，最高9千万</td></tr>"   
   response.write "<tr><td class=tablebody2>自动清理事件：</td><td width=15% class=tablebody1><input type=radio name=logbz value=1" 
   if log_set(0)="1" then response.write " checked"
   response.write ">开启<input type=radio name=logbz value=0"
   if log_set(0)="0" then response.write " checked"
   response.write " >关闭</td>"
   response.write "<td class=tablebody1 colspan=2>事件上限数：<input name=lognum size=18 maxlength=20 value="&log_set(1)&"> 设定囚犯被抢劫和探监事件自动清理功能</td></tr>"   
   response.write "<tr><td class=tablebody2 rowspan=2>抢劫功能设置：</td>"
   response.write "<td width=15% class=tablebody1>开放抢劫囚犯功能</td>"
   response.write "<td class=tablebody1><input type=radio name=qjbz value=1 "
   if qj_set(0)="1" then response.write " checked"
   response.write ">开放<input type=radio name=qjbz value=0 "
   if qj_set(0)="0" then response.write " checked"
   response.write ">关闭</td>"
   response.write "<td class=tablebody1>抢劫囚犯者是否扣其魅力值？&nbsp;&nbsp;<input type=radio name=qjmlbz value=1 "
   if qj_set(1)="1" then response.write " checked"
   response.write ">扣除<input type=radio name=qjmlbz value=0 "
   if qj_set(1)="0" then response.write " checked"
   response.write ">不扣　  魅力值<input name=qjml size=4 maxlength=4 value="&qj_set(2)&"> </td></tr>"
   response.write "<tr><td class=tablebody1>显示抢劫者姓名？</td>"
   response.write "<td class=tablebody1><input type=radio name=qjxm value=1 "
   if qj_set(3)="1" then response.write " checked"
   response.write ">显示<input type=radio name=qjxm value=0 "
   if qj_set(3)="0" then response.write " checked"
   response.write ">关闭</td>"
   response.write "<td class=tablebody1>是否限定20分钟抢劫次数限制？<input type=radio name=qjtime value=1 "
   if qj_set(4)="1" then response.write " checked"
   response.write ">限制<input type=radio name=qjtime value=0 "
   if qj_set(4)="0" then response.write " checked"
   response.write ">不限"
   response.write "</td></tr>"
response.write "<tr><td class=tablebody2 colspan=4><div align=center><input type=submit name=Submit value='保存修改'></div></td></tr>"
   response.write "</table></form></CENTER></td></tr></table>"
  end if 
end if
call footer() 
%>
