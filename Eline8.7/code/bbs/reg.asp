<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/email.asp"-->
<!--#include file="md5.asp"-->
<%
'世纪江湖动网论坛插件 回首当年 www.happyjh.com
response.redirect "../yamen/read.htm"

'=========================================================
' File: reg.asp
' Version:5.0
' Date: 2002-12-3
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
response.buffer=true
Dim SendMail
call nav()
if Cint(forum_setting(37))=0 then
	stats="论坛错误信息"
	Errmsg="<br><li>本论坛暂时停止注册。"
	call head_var(2,0,"","")
	call dvbbs_error()
else
	if request("action")="apply" then
		stats="填写资料"
		call head_var(0,0,"论坛注册","reg.asp")
		call reg_2()
	elseif request("action")="save" then
		stats="提交注册"
		call head_var(0,0,"论坛注册","reg.asp")
		call reg_3()
		if founderr then call dvbbs_error()
	else
		stats="注册协议"
		call head_var(0,0,"论坛注册","reg.asp")
		call reg_1()
	end if
end if
call activeonline()
call footer()

sub reg_1()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr><th align=center><form action="reg.asp?action=apply" method=post>服务条款和声明</td></tr>
    <tr><td class=tablebody1 align=left>
<!--#include file="inc/reg_txt.asp"-->
	</td></tr>
    <tr><td align=center class=tablebody2><input type=submit value=我同意></td></form></tr>
</table>
<%
end sub

sub reg_2()
%>
<FORM name=theForm action=reg.asp?action=save method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th colSpan=2 height=24>新用户注册</TD>
</TR>
<TR> 
<TD width=40% class=tablebody1><B>用户名</B>：<BR>注册用户名长度限制为<%=forum_setting(40)%>－<%=forum_setting(41)%>字节</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength="<%=forum_setting(41)%>" size=30 name=name></TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>性别</B>：<BR>请选择您的性别</font></TD>
<TD width=60%  class=tablebody1> <INPUT type=radio CHECKED value=1 name=sex>
<IMG  src=pic/Male.gif align=absMiddle>男孩 &nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT type=radio value=0 name=sex>
<IMG  src=pic/Female.gif align=absMiddle>女孩</font></TD>
</TR>
<%if cint(Forum_Setting(23))<>1 then%>
<TR> 
<TD width=40% class=tablebody1><B>密码(至少6位)</B>：<BR>
请输入密码，区分大小写。<BR>
请不要使用任何类似 '*'、' ' 或 HTML 字符
</TD>
<TD width=60% class=tablebody1>
<INPUT type=password maxLength=16 size=30 name=psw>
</TD>
</TR>
<TR> 
<TD width=40% class=tablebody1><B>密码(至少6位)</B>：<BR>请再输一遍确认</TD>
<TD class=tablebody1> 
<INPUT type=password maxLength=16 size=30 name=pswc>
</TD>
</TR>
<%end if%>
<TR> 
<TD width=40%  class=tablebody1><B>密码问题</B>：<BR>忘记密码的提示问题</TD>
<TD class=tablebody1> 
<INPUT type=text size=30 name=quesion>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>问题答案</B>：<BR>忘记密码的提示问题答案，用于取回论坛密码</TD>
<TD class=tablebody1> 
<INPUT type=text size=30 name=answer>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>Email地址</B>：<BR>请输入有效的邮件地址，这将使您能用到论坛中的所有功能</font></TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=50 size=30 name=e_mail><input type=button value='检测帐号' name=Button onclick=gopreview()></TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1 id=adv style="DISPLAY: none">
<TBODY> 
<TR align=middle> 
<Th colSpan=2 height=24 align=left>填写详细资料</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>头像</B>：<BR>选择的头像将出现在您的资料和发表的帖子中，您也可以选择下面的自定义头像</TD>
<TD width=60%  class=tablebody1> 
<select name=face size=1 onChange="document.images['face'].src=options[selectedIndex].value;" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
<%for i=0 to Forum_userfaceNum%>
<option value="<%=Forum_info(11)&Forum_userface(i)%>"><%=Forum_userface(i)%></option>
<%next%>
</select>
&nbsp;<img id=face src=<%=Forum_info(11)&Forum_userface(0)%>>&nbsp;<a href=allface.asp target=_blank>查看所有头像</a>
</TR>
<%if Cint(forum_setting(54))=0 then%>
<TR> 
<TD width=40% valign=top class=tablebody1><B>自定义头像</B>：<br>如果图像位置中有连接图片将以自定义的为主</TD>
<TD width=60%  class=tablebody1>
<%if cint(Forum_Setting(7))=1 then%>
<iframe name=ad frameborder=0 width=300 height=40 scrolling=no src=reg_upload.asp></iframe> 
<br>
<%end if%>
图像位置： 
<input type=TEXT name=myface size=20 maxlength=100>
&nbsp;完整Url地址<br>
宽&nbsp;&nbsp;&nbsp;&nbsp;度： 
<input type=TEXT name=width size=3 maxlength=3 value="<%=forum_setting(38)%>">
0---<%=forum_setting(57)%>的整数<br>
高&nbsp;&nbsp;&nbsp;&nbsp;度： 
<input type=TEXT name=height size=3 maxlength=3 value="<%=forum_setting(39)%>">
0---<%=forum_setting(57)%>的整数<br>
</TD>
</TR>
<%end if%>
      <tr bgcolor="&Forum_body(4)&">    
        <td width=40%  class=tablebody1><B>生日</B><BR>如不想填写，请全部留空</td>   
        <td width=60%  class=tablebody1>
<select name=birthyear>
<option value=></option>
<%for i=1901 to 2002%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
                   年 
<select name=birthmonth>
<option value=></option>
<%for i=1 to 12%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
                   月 
<select name=birthday>
<option value=></option>
<%for i=1 to 31%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
日
        </td>   
      </tr>
<tr> 
<td width=40%  class=tablebody1><B>回复提示</B>：<BR>当您发表的帖子有人回复时，使用论坛信息通知您。</td>
<td width=60%  class=tablebody1>
<input type=radio name=showRe value=1 checked>
提示我
<input type=radio name=showRe value=0>
不提示
</tr>
<%if cint(Forum_Setting(32))=1 then%>
<TR> 
<TD width=40% class=tablebody1><B>门派</B>：<BR>您可以自由选择要加入的门派</TD>
<TD width=60% class=tablebody1> 
<select name=groupname>
<%
set rs=conn.execute("select * from GroupName")
if rs.eof and rs.bof then
	response.write "<option value=无门无派>无门无派</option>"
else
	do while not rs.eof
	response.write "<option value="&rs("Groupname")&">"&rs("GroupName")&"</option>"
	rs.movenext
	loop
end if
rs.close
set rs=nothing
%>
</select>
</TD>
</TR>
<%end if%>
<TR> 
<TD width=40%  class=tablebody1><B>OICQ号码</B>：<BR>填写您的QQ地址，方便与他人的联系</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=20 size=44 name=OICQ>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>ICQ号码</B>：<BR>填写您的ICQ地址，方便与他人的联系</font></TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=20 size=44 name=ICQ>
</TD>
</TR>
<TR > 
<TD width=40%  class=tablebody1><B>MSN</B>：<BR>填写您的MSN地址，方便与他人的联系</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=70 size=44 name=msn>
</TD>
</TR>
<%if Cint(forum_setting(42))=1 then%>
<TR> 
<TD width=40%  class=tablebody1><B>签名</B>：<BR>最多300字节<BR>
文字将出现在您发表的文章的结尾处。体现您的个性。 </TD>
<TD width=60%  class=tablebody1> 
<TEXTAREA name=Signature rows=5 wrap=PHYSICAL cols=60></TEXTAREA>
</TD>
</TR>
<%end if%>
<tr>    
<td width=40%  class=tablebody1><B>选择Cookie的保留时间</B>：<BR>登陆论坛信息保留时间，在这个时间内重复登陆论坛不需要重新登陆</font></td>  
<td width=60%  class=tablebody1>    
			              <input type=radio name=usercookies value=1 checked>
              <font color=red>1天</font> 
              <input type=radio name=usercookies value=2>
    1个月
    <input type=radio name=usercookies value=3>
    1年
              <input type=radio name=usercookies value=0>
    不保留 </td>   
      </tr>
<tr>    
<td width=40%  class=tablebody1><B>是否开放您的基本资料</B>：<BR>开放后别人可以看到您的性别、Email、QQ等信息</td>  
<td width=60%  class=tablebody1>    
			              <input type=radio name=setuserinfo value=1 checked>
              开放              <input type=radio name=setuserinfo value=0>
    不开放 </td>   
      </tr>
<tr>    
<td width=40%  class=tablebody1><B>是否开放您的真实资料</B>：<BR>开放后别人可以看到您的真实姓名、联系方式等信息</td>
  <td width=60%  class=tablebody1>    
			
              <input type=radio name=setusertrue value=1>
              开放              <input type=radio name=setusertrue value=0 checked>
    不开放</td>   
      </tr>
<tr>
<th height=25 align=left valign=middle colspan=2><b>&nbsp;个人真实信息</b>（以下内容建议填写）</th>
</tr>
<tr>
<td height=2 valign=top  class=tablebody1 width=40% > 　<b>真实姓名：</b>
<input type=text name=realname size=18>
</td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>性　格： &nbsp; </b>
<br>
<% dim KidneyType,theKidney
  KidneyType="多重性格,乐天达观,成熟稳重,幼稚调皮,温柔体贴,活泼可爱,普普通通,内向害羞,外向开朗,心地善良,聪明伶俐,善解人意,风趣幽默,思想开放,积极进取,小心谨慎,郁郁寡欢,正义正直,悲观失意,好吃懒做,处事洒脱,疑神疑鬼,患得患失,异想天开,多愁善感,淡泊名利,见利忘义,瞻前顾后,循规蹈矩,热心助人,快言快语,少言寡语,爱管闲事,追求刺激,豪放不羁,狡猾多变,贪小便宜,见异思迁,情绪多变,水性扬花,重色轻友,胆小怕事,积极负责,勇敢正义,聪明好学,实事求是,务实实际,老实巴交,圆滑老练,脾气暴躁,慢条斯理,诚实坦白,婆婆妈妈,急性子"
  theKidney=split(KidneyType, ",")
	   for i = 0 to ubound(theKidney)	
	   response.write "<input type=""checkbox"" name=""character"" value="""&trim(theKidney(i))&""" "	 
	   response.write ">"&trim(theKidney(i))&" "
	   if ((i+1) mod 5)=0 then response.write "<br>"  '每行显示六个性格进行换行
	   next   
  %></td>
</tr>
<tr>
<td class=tablebody1><b>个人简介： &nbsp;</b><br>
<textarea name=personal rows=6 cols=90% ></textarea>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>国　　家：</b>
<b>
<input type=text name=country size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>联系电话：</b>
<b>
<input type=text name=userphone size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>通信地址：</b>
<b>
<input type=text name=address size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>省　　份：</b>
<input type=text name=province size=18>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>城　　市：
</b>
<input type=text name=city size=18>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>生　　肖：
</b>
<select size=1 name=shengxiao>
<option></option>
<option value=鼠>鼠</option>
<option value=牛>牛</option>
<option value=虎>虎</option>
<option value=兔>兔</option>
<option value=龙>龙</option>
<option value=蛇>蛇</option>
<option value=马>马</option>
<option value=羊>羊</option>
<option value=猴>猴</option>
<option value=鸡>鸡</option>
<option value=狗>狗</option>
<option value=猪>猪</option>
</select>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>血　　型：</b>
<select size=1 name=blood>
<option selected></option>
<option value=A>A</option>
<option value=B>B</option>
<option value=AB>AB</option>
<option value=O>O</option>
<option value=其他>其他</option>
</select>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >　<b>信　　仰：</b>
<select size=1 name=belief>
<option selected></option>
<option value=佛教>佛教</option>
<option value=道教>道教</option>
<option value=基督教>基督教</option>
<option value=天主教>天主教</option>
<option value=回教>回教</option>
<option value=无神论者>无神论者</option>
<option value=共产主义者>共产主义者</option>
<option value=其他>其他</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >　<b>职　　业： </b>
<select name=occupation>
<option selected> </option>
<option value=财会/金融>财会/金融</option>
<option value=工程师>工程师</option>
<option value=顾问>顾问</option>
<option value=计算机相关行业>计算机相关行业</option>
<option value=家庭主妇>家庭主妇</option>
<option value=教育/培训>教育/培训</option>
<option value=客户服务/支持>客户服务/支持</option>
<option value=零售商/手工工人>零售商/手工工人</option>
<option value=退休>退休</option>
<option value=无职业>无职业</option>
<option value=销售/市场/广告>销售/市场/广告</option>
<option value=学生>学生</option>
<option value=研究和开发>研究和开发</option>
<option value=一般管理/监督>一般管理/监督</option>
<option value=政府/军队>政府/军队</option>
<option value=执行官/高级管理>执行官/高级管理</option>
<option value=制造/生产/操作>制造/生产/操作</option>
<option value=专业人员>专业人员</option>
<option value=自雇/业主>自雇/业主</option>
<option value=其他>其他</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >　<b>婚姻状况：</b>
<select size=1 name=marital>
<option selected></option>
<option value=未婚>未婚</option>
<option value=已婚>已婚</option>
<option value=离异>离异</option>
<option value=丧偶>丧偶</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >　<b>最高学历：</b>
<select size=1 name=education>
<option selected></option>
<option value=小学>小学</option>
<option value=初中>初中</option>
<option value=高中>高中</option>
<option value=大学>大学</option>
<option value=硕士>硕士</option>
<option value=博士>博士</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >　<b>毕业院校：</b>
<input type=text name=college size=18></td>
</tr>
</TBODY> 
</TABLE>
</td></tr></table>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center>
<tr>
<td width=50% height=24><INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()><span id=advance>显示高级用户设置选项</a></span> </td>
<td width=50% ><input type=submit value="注 册" name=Submit><input type=reset value="清 除" name=Submit2></td>
</tr></table>
</form>
<form name=preview action=chkreg.asp method=post target=preview_page>
<input type=hidden name=username value=><input type=hidden name=email value=>
</form>
<script>
function gopreview()
{
document.forms[1].username.value=document.forms[0].name.value;
document.forms[1].email.value=document.forms[0].e_mail.value;
var popupWin = window.open('preview.asp', 'preview_page', 'scrollbars=yes,width=500,height=300');
document.forms[1].submit()
}
function showadv(){
if (document.theForm.advshow.checked == true) {
	adv.style.display = "";
	advance.innerText="关闭高级用户设置选项"
}else{
	adv.style.display = "none";
	advance.innerText="显示高级用户设置选项"
}
}
</script>
<%
end sub

sub reg_3()
dim username,sex,pass1,pass2,password
dim useremail,face,width,height
dim oicq,sign,showRe,birthday
dim mailbody,sendmsg,rndnum,num1
dim quesion,answer,topic
dim userinfo,usersetting
dim userclass
if not isnull(session("regtime")) or cint(Forum_Setting(22))>0 then
	if DateDiff("s",session("regtime"),Now())<cint(Forum_Setting(22)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>本论坛限制每次注册距离时间为"&Forum_Setting(22)&"秒，请稍后注册。"
   		FoundErr=True
	end if
end if

if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
   	FoundErr=True
end if
if request("name")="" or strLength(request("name"))>Cint(forum_setting(41)) or strLength(request("name"))<Cint(forum_setting(40)) then
	errmsg=errmsg+"<br>"+"<li>请输入您的用户名(长度不能大于"&forum_setting(41)&"小于"&forum_setting(40)&")。"
	founderr=true
else
	username=trim(request("name"))
end if
if Instr(request("name"),"=")>0 or Instr(request("name"),"%")>0 or Instr(request("name"),chr(32))>0 or Instr(request("name"),"?")>0 or Instr(request("name"),"&")>0 or Instr(request("name"),";")>0 or Instr(request("name"),",")>0 or Instr(request("name"),"'")>0 or Instr(request("name"),",")>0 or Instr(request("name"),chr(34))>0 or Instr(request("name"),chr(9))>0 or Instr(request("name"),"")>0 or Instr(request("name"),"$")>0 then
	errmsg=errmsg+"<br>"+"<li>用户名中含有非法字符。"
	founderr=true
else
	username=trim(request("name"))
end if

splitwords=split(splitwords,",")
for i = 0 to ubound(splitwords)
	if instr(username,splitwords(i))>0 then
		errmsg=errmsg+"<br>"+"<li>您输入的用户名包含系统禁止注册字符。"
		founderr=true
		exit for
	end if
next

if request("sex")=0 or request("sex")=1 then
	sex=request("sex")
else
	sex=1
end if
	
if request("showRe")=0 or request("showRe")=1 then
	showRe=request("showRe")
else
	showRe=1
end if	
if cint(Forum_Setting(23))=1 then
	Randomize
	Do While Len(rndnum)<8
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	password=md5(rndnum)
else
	if request("psw")="" or len(request("psw"))>10 or len(request("psw"))<6 then
		errmsg=errmsg+"<br>"+"<li>请输入您的密码(长度不能大于10小于6)。"
		founderr=true
	else
	pass1=request("psw")
	end if
	if request("pswc")="" or strLength(request("pswc"))>10 or len(request("pswc"))<6 then
		errmsg=errmsg+"<br>"+"<li>请输入确认密码(长度不能大于10小于6)。"
		founderr=true
	else
		pass2=request("pswc")
	end if
	if pass1<>pass2 then
		errmsg=errmsg+"<br>"+"<li>您输入的密码和确认密码不一致。"
		founderr=true
	else
		password=md5(pass2)
	end if
end if

if request("quesion")="" then
  	errmsg=errmsg+"<br>"+"<li>请输入密码提示问题。"
	founderr=true
else
	quesion=request("quesion")
end if
if request("answer")="" then
  	errmsg=errmsg+"<br>"+"<li>请输入密码提示问题答案。"
	founderr=true
elseif request("answer")=request("oldanswer") then
	answer=request("answer")
else
	answer=md5(request("answer"))
end if

if IsValidEmail(trim(request("e_mail")))=false then
	errmsg=errmsg+"<br>"+"<li>您的Email有错误。"
   	founderr=true
else
	if not isnull(forum_setting(52)) and forum_setting(52)<>"" and forum_setting(52)<>"0" then
		Dim SplitUserEmail
		SplitUserEmail=split(forum_setting(52),"|")
		For i=0 to ubound(SplitUserEmail)
			if instr(request("e_mail"),SplitUserEmail(i))>0 then
			errmsg=errmsg+"<br>"+"<li>您填写的Email地址含有系统禁止字符。"
   			founderr=true
			exit for
			end if
		Next
	end if
	useremail=checkStr((request("e_mail")))
end if
if request.form("myface")<>"" and Cint(forum_setting(54))=0 then
	if request("width")="" or request("height")="" then
		errmsg=errmsg+"<br>"+"<li>请输入图片的宽度和高度。"
		founderr=true
	elseif not isInteger(request("width")) or not isInteger(request("height")) then
		errmsg=errmsg+"<br>"+"<li>您输入的字符不合法。"
		founderr=true
	elseif Cint(request("width"))>Cint(forum_setting(57)) then
		errmsg=errmsg+"<br>"+"<li>您输入的图片宽度不符合标准。"
		founderr=true
	elseif Cint(request("height"))>Cint(forum_setting(57)) then
		errmsg=errmsg+"<br>"+"<li>您输入的图片高度不符合标准。"
		founderr=true
	else
		if Cint(forum_setting(55))=0 then
			if instr(lcase(request("myface")),"http://")>0 or instr(lcase(request("myface")),"www.")>0 then
				errmsg=errmsg+"<br>"+"<li>本论坛设置了不允许连接外部网址做头像。"
				founderr=true
			end if
		end if
		face=request("myface")
	end if
else
	if request("face")<>"" then
		face=request("face")
	end if
end if
		width=request("width")
		height=request("height")
if request("oicq")<>"" then
	if not isnumeric(request("oicq")) or len(request("oicq"))>10 then
		errmsg=errmsg+"<br>"+"<li>Oicq号码只能是4-10位数字，您可以选择不输入。"
		founderr=true
	end if
end if
if request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" then
	birthday=""
else
	birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
	if not isdate(birthday) then birthday=""
end if
userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")
if founderr then exit sub
dim titlepic
set rs=conn.execute("select usertitle,titlepic from usertitle where not minarticle=-1 order by minarticle")
userclass=rs(0)
titlepic=rs(1)
set rs=server.createobject("adodb.recordset")
if cint(Forum_Setting(24))=1 then
sql="select * from [user] where username='"&username&"' or useremail='"&useremail&"'"
else
sql="select * from [user] where username='"&username&"'"
end if
rs.open sql,conn,1,3
if not rs.eof and not rs.bof then
	if cint(Forum_Setting(24))=1 then
		errmsg=errmsg+"<br>"+"<li>您输入的用户名已经被注册或者已经有用户使用了您填写的电子邮件地址。"
		founderr=true
		exit sub
	else
		errmsg=errmsg+"<br>"+"<li>您输入的用户名已经被注册。"
		founderr=true
		exit sub
	end if
else
	rs.addnew
	rs("username")=username
	rs("userpassword")=password
	rs("useremail")=useremail
	rs("userclass")=userclass
	rs("titlepic")=titlepic
	rs("quesion")=quesion
	rs("answer")=answer
	
	if request("Signature")<>"" then
		rs("sign")=trim(request("Signature"))
	end if
	if request("oicq")<>"" then
		rs("oicq")=request("oicq")
	end if
	if request("icq")<>"" then
		rs("icq")=request("icq")
	end if
	if request("msn")<>"" then
		rs("msn")=request("msn")
	end if
	Rs("article")=0
	if cint(Forum_Setting(25))=1 then
		rs("usergroupid")=5
	else
       	Rs("usergroupid")=4
	end if
	rs("lockuser")=0
	Rs("sex")=sex
	Rs("showRe")=showRe
	if birthday<>"" then
		rs("birthday")=birthday
	end if
	rs("UserGroup")=request("GroupName")
	Rs("addDate")=NOW()
	if request.form("myface")<>"" then
		rs("face")=face
	else
		rs("face")=face
	end if
	Rs("width")=width
	Rs("height")=height
	rs("logins")=1
	Rs("lastlogin")=NOW()
	rs("userWealth")=Forum_user(0)
	rs("userEP")=Forum_user(5)
	rs("usercP")=Forum_user(10)
	rs("userinfo")=userinfo
	rs("usersetting")=usersetting
	rs("bbstype")=tempid
	rs.update
	conn.execute("update config set usernum=usernum+1,lastuser='"&username&"'")
end if
rs.close
set rs=nothing


set rs=conn.execute("select top 1 userid,face from [user] order by userid desc")
userid=rs(0)

'******************
'对上传头象进行过滤与改名
if Cint(Forum_Setting(7))=1 then 
	on error resume next
	dim objFSO,facename
	dim upface,newfilename
	facename=rs(1)
	if instr(facename,"uploadFace/") then
	facename=split(facename,"/")
	upface="uploadFace/"&facename(ubound(facename))
	newfilename="uploadFace/"&userid&"_"&facename(ubound(facename))
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if	objFSO.fileExists(Server.MapPath(upface)) then
	objFSO.movefile ""&Server.MapPath(upface)&"",""&Server.MapPath(newfilename)&""
	if Err.Number = 0 then
	conn.execute("update [user] set face='"&newfilename&"' where userid="&userid)
	end if
	end if
	set objFSO=nothing
	end if
end if

'对上传头象进行过滤与改名结束
'****************
rs.close
set rs=nothing

if Forum_Setting(47)=1 then
	on error resume next
	'发送注册邮件
	dim getpass
	topic="您在" & Forum_info(0) & "的注册资料"
	if cint(Forum_Setting(23))=1 then
		getpass=htmlencode(rndnum)
	else
		getpass=htmlencode(request("psw"))
	end if
%>
<!--#include file="inc/email_txt.asp"-->
<%
	select case Cint(Forum_Setting(2))
	case 0
		sendmsg="<li>系统未开启邮件功能，请记住您的注册信息。</li>"
	case 1
		call jmail(useremail,topic,mailbody)
	case 2
		call Cdonts(useremail,topic,mailbody)
	case 3
		call aspemail(useremail,topic,mailbody)
	case else
		sendmsg="<li>系统未开启邮件功能，请记住您的注册信息。</li>"
	end select
	if SendMail="OK" then
		if cint(Forum_Setting(23))=1 then
		sendmsg="<li>您的注册信息和密码已经发往您的邮箱，请使用系统给您的密码登陆。</li>"
		else
		sendmsg="<li>您的注册信息已经发往您的邮箱，请注意查收。</li>"
		end if
	else
		sendmsg="<li>由于系统错误，给您发送的注册资料未成功。</li>"
	end if
	'response.write mailbody
end if

if Forum_Setting(46)=1 then
	'发送注册短信
	dim sender,title,body
	sender=Forum_info(0)
	title=Forum_info(0)&"欢迎您的到来"
%>
<!--#include file="inc/sms_txt.asp"-->
<%
	'response.write body
	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
	conn.Execute(sql)
end if

if cint(Forum_Setting(23))=1 or cint(Forum_Setting(25))=1 then

else
	if founduser then
		Response.Cookies("aspsky").path=cookiepath
		Response.Cookies("aspsky")("username")=""
		Response.Cookies("aspsky")("password")=""
		Response.Cookies("aspsky")("userclass")=""
		Response.Cookies("aspsky")("userid")=""
		Response.Cookies("aspsky")("userhidden")=""
		Response.Cookies("aspsky")("usercookies")=""
		conn.execute("delete from online where username='"&membername&"'")
	end if
	select case request("usercookies")
 	case 0
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 1
	    Response.Cookies("aspsky").Expires=Date+1
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 2
		Response.Cookies("aspsky").Expires=Date+31
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 3
	    Response.Cookies("aspsky").Expires=Date+365
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
	end select
	Response.Cookies("aspsky")("username") = username
	Response.Cookies("aspsky")("password") = password
	Response.Cookies("aspsky")("userclass") = userclass
	Response.Cookies("aspsky")("userid") = userid
	Response.Cookies("aspsky")("userhidden") = 2
	Response.Cookies("aspsky").path=cookiepath
end if
session("regtime")=now()
%>
<meta HTTP-EQUIV=REFRESH CONTENT='2; URL=index.asp'>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>注册成功：<%=Forum_info(0)%>欢迎您的到来</th>
</tr>
<tr><td class=tablebody1><br>
<ul><%=sendmsg%><li><a href="index.asp">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>