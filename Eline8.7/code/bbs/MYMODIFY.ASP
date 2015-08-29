<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<%
'=========================================================
' File: mymodify.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
response.buffer=true
stats="修改基本资料"
dim sex
dim showre
dim face,width,height
dim birthday
dim usercookies
dim msgFlag
call nav()

if Cint(GroupSetting(16))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛修改资料的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您还没有登录，请登录后进行修改。"
	founderr=true
end if

if founderr then
  errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
	if request("action")="updat" then
		if lcase(membername)="test" then
		  errmsg=errmsg+"<br>"+"<li>测试用户不能修改个人资料。"
			call dvbbs_error()
		else
			call update()
			if founderr then
				call dvbbs_error()
			else
				sucmsg="<li>更改资料成功！"
				call dvbbs_suc()
			end if
		end if
	else
		call userinfo()
		if founderr then call dvbbs_error()
	end if
end if

call activeonline()
call footer()

sub userinfo()
dim CanUseTitle,CanUseTitle1,CanUseTitle2
CanUseTitle=false
CanUseTitle1=false
CanUseTitle2=false
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
else
	if Cint(Forum_Setting(6))=1 then
		CanUseTitle=true
	end if
	if CanUseTitle and Cint(forum_setting(60))>0 and rs("article")>Cint(forum_setting(60)) then
		CanUseTitle1=true
	elseif CanUseTitle and Cint(forum_setting(60))=0 then
		CanUseTitle1=true
	else
		CanUseTitle1=false
	end if
	if CanUseTitle and Cint(forum_setting(61))>0 and DateDiff("d",rs("AddDate"),Now())>Cint(forum_setting(61)) then
		CanUseTitle2=true
	elseif CanUseTitle and Cint(forum_setting(61))=0 then
		CanUseTitle2=true
	else
		CanUseTitle2=false
	end if
	if CanUseTitle and Cint(forum_setting(62))=1 then
		if CanUseTitle1 and CanUseTitle2 then
			CanUseTitle=true
		else
			CanUseTitle=false
		end if
	elseif CanUseTitle and (CanUseTitle1 or CanUseTitle2) then
		CanUseTitle=true
	else
		CanUseTitle=false
	end if
%>

<!--#include file="z_controlpanel.asp"-->
<br>
<form action="mymodify.asp?action=updat&username=<%=htmlencode(membername)%>" method=POST name="theForm">
<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
<tr> 
      <th colspan="2" width="100%">基本设置选项
      </th>
    </tr> 
<%if CanUseTitle then%>
      <tr> 
        <td width="40%" class=tablebody1><B>头&nbsp;&nbsp;&nbsp;&nbsp;衔</B>：</td>   
        <td width="60%" class=tablebody1> <input type="text" name="title" value="<%if trim(rs("title"))<>"" then%><%=htmlencode(RS("title"))%><%end if%>" size=15 maxlength=13> </td>   
      </tr>   
<%end if%>
      <tr>    
        <td width="40%" class=tablebody1><B>性&nbsp;&nbsp;&nbsp;&nbsp;别</B>：<BR>请选择您的性别</td>   
        <td width="60%" class=tablebody1><%if rs("sex")=1 then%>酷哥<%elseif rs("sex")=0 then%>靓妹<%end if%></td>
</td>   
      </tr>
      <tr>    
        <td width="40%" class=tablebody1><B>生&nbsp;&nbsp;&nbsp;&nbsp;日</B><BR>如不想填写，请全部留空</td>   
        <td width="60%" class=tablebody1>    
				  <select name=birthyear>
				  <option value=""></option>
				  <%for i=1901 to 2000%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(year(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    年 
				  <select name=birthmonth>
				  <option value=""></option>
				  <%for i=1 to 12%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(month(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    月 
				  <select name=birthday>
				  <option value=""></option>
				  <%for i=1 to 31%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(day(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    日</font>
        </td>   
      </tr>
      <tr>    
        <td width="40%" valign=top class=tablebody1><B>头&nbsp;&nbsp;&nbsp;&nbsp;像</B>：(<font color=<%=forum_body(8)%>>修改头像将收取费用<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(22)
					else
						response.write forum_user(23)
					end if
					%>元，您现在有现金<%=mymoney%>元</font>)<BR>您可以选择一个个性图片，当你发表时将显示在您的名字下方。<%
        	if (Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0 then
						if Cint(Forum_Setting(7))=1 then
							%><BR>如果你填写了下面的自定义头像部分，那么你的头像以自定义的为准。否则，请你留空自定义头像的所有栏目！<BR><BR><B>关于自定义头像</B>：(<font color=<%=forum_body(8)%>>上传头像将收取费用<%
							if isnull(myvip) or myvip<>1 then
								response.write forum_user(34)
							else
								response.write forum_user(35)
							end if
							%>元/KB，您现在有现金<%=mymoney%>元</font>)<BR>你也可以在这里给出你自定义头像的 URL 地址，头像的高度和宽度(像素)。 如果不想要自定义头像，请将相应栏目栏目全部留空！<%
						end if
					end if
					%></td>   
        <td width="60%" class=tablebody1>
 <table border="0" cellpadding="0" cellspacing="1" width="98%" >
            <tr>
              <td width="*"  class=tablebody1>论坛头像选择 
<select name=face style="font-family:arial" size=1 onChange="document.images['face'].src=options[selectedIndex].value; <%if (Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0 then%>parent.document.all.theForm.myface.value=options[selectedIndex].value; parent.document.all.theForm.width.value=<%=forum_setting(38)%>;parent.document.all.theForm.height.value=<%=forum_setting(39)%>;<%end if%>document.images['face'].width=<%=forum_setting(38)%>;document.images['face'].height=<%=forum_setting(39)%>;"> 
<%
	dim userregface
	for i=0 to Forum_userfaceNum-1
	userregface=Forum_info(11)&Forum_userface(i)
	if trim(lcase(rs("face"))) = trim(lcase(userregface)) then
%> 
   <option value='<%=Forum_info(11)&Forum_userface(i)%>' selected><%=Forum_userface(i)%></option> 
     <% else %> 
   <option value='<%=Forum_info(11)&Forum_userface(i)%>'><%=Forum_userface(i)%></option> 
<%
	end if 
	next
%> 
                </select><a href="allface.asp" target="_blank" title="头像预览列表">[<font color=red>头像预览</font>]</a></td>
              <td width="180"  rowspan="2"  valign=top align=center>
<img id=face src="<%=rs("face")%>" width=<%=rs("width")%> height=<%=rs("height")%>>
</td>
            </tr>
            <tr>
              <td width="*" class=tablebody1>
<%if (Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0 then%>
<%if Cint(Forum_Setting(7))=1 then %> 
<iframe name="ad" frameborder=0 width=300 height=26 scrolling=no src=reg_upload.asp></iframe>
<% end if %>
<br>自定义头像地址：<input type="TEXT" name="myface" size=30 maxlength=100 value="<%=rs("face")%>" <%if not ((Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0) or Cint(Forum_Setting(7))=0 then%>disabled<%end if%>> 
<br>宽&nbsp;&nbsp;&nbsp;&nbsp;度：<input type="TEXT" name="width" size=3 maxlength=3 value="<%=rs("width")%>" <%if not ((Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0) or Cint(Forum_Setting(7))=0 then%>disabled<%end if%>>&nbsp;0---<%=forum_setting(57)%>的整数
<br>高&nbsp;&nbsp;&nbsp;&nbsp;度：<input type="TEXT" name="height" size=3 maxlength=3 value="<%=rs("height")%>" <%if not ((Cint(forum_setting(54))>0 and rs("article")>Cint(forum_setting(54))) or Cint(forum_setting(54))=0) or Cint(Forum_Setting(7))=0 then%>disabled<%end if%>>&nbsp;0---<%=forum_setting(58)%>的整数
<br>
<%END IF%>
</td>
            </tr>
          </table>
        </td>   
      </tr>	  
      <tr>    
        <td width="40%" class=tablebody1><B>个人照片</B>：<BR>如果您有照片在网上，请输入网页地址。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="userphoto" value="<%if trim(rs("userphoto"))<>"" then%><%=htmlencode(rs("userphoto"))%><%end if%>" size=30 maxlength=100>   
        </td>   
      </tr>   
<%if cint(Forum_Setting(32))=1 then%>
      <tr>    
        <td width="40%" class=tablebody1><B>门&nbsp;&nbsp;&nbsp;&nbsp;派</B>：<BR>如果您有兴趣，可加入以下门派</td>   
        <td width="60%" class=tablebody1><%=rs("UserGroup")%> <font color=gray>[如果想改入其他门派,请去论坛的<A href="z_menpai.asp">江湖门派</a>]</font></td>   
      </tr>
<%end if%>
<%if Cint(forum_setting(42))=1 then%>
      <tr>    
        <td valign=top width="40%" class=tablebody1><B>签&nbsp;&nbsp;&nbsp;&nbsp;名</B>：(<font color=<%=forum_body(8)%>>修改签名将收取费用<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(24)
					else
						response.write forum_user(25)
					end if
				%>元，您现在有现金<%=mymoney%>元</font>)<BR>不能超过 250 个字符<br>文字将出现在您发表的文章的结尾处。</td>
        <td width="60%" class=tablebody1><textarea name="Signature" rows=5 cols=60 wrap=PHYSICAL><%
        	if rs("sign")<>"" then
						dim signtrue
						signtrue=replace(rs("sign"),"<BR>",chr(13))
						signtrue=replace(signtrue,"&nbsp;"," ")
						%><%=signtrue%><%
					end if%></textarea>   
        </td>   
      </tr>
<%end if%>
<tr> 
      <th colspan="2" width="100%">其他选项
      </th>
    </tr> 
 <tr>    
        <td width="40%" class=tablebody1><B>有回帖时是否提示您</B>：<BR>当您的帖子有人回复后有信息通知您</td>   
        <td width="60%" class=tablebody1>    
			<input type="radio" name="showRe" value="1" <% if rs("showRe") then response.write "checked" %> >
			提示我
			<input type="radio" name="showRe" value="0" <% if not rs("showRe") then response.write "checked" %> >
			不提示
        </td>   
      </tr>	  
 <tr>    
        <td width="40%" class=tablebody1><B>接收短消息的对象</B>：<BR>选择哪些人可以向您发送短消息。</td>   
        <td width="60%" class=tablebody1>    
			<input type="radio" name="msgFlag" value="0" <% if isnull(rs("msgFlag")) or rs("msgFlag")=0 then response.write "checked" %> >
			所有人
			<input type="radio" name="msgFlag" value="1" <% if not isnull(rs("msgFlag")) and rs("msgFlag")=1 then response.write "checked" %> >
			仅好友
			<input type="radio" name="msgFlag" value="9" <% if not isnull(rs("msgFlag")) and rs("msgFlag")=9 then response.write "checked" %> >
			不接收
        </td>   
      </tr>	  
      <tr>    
        <td width="40%" class=tablebody1><B>选择Cookie的保留时间</B>：</td>   
        <td width="60%" class=tablebody1>    
			
    <input type="radio" name="usercookies" value="1" <% if request.cookies("aspsky")("usercookies")=1 then response.write "checked" %> >
    <font color=red>1天</font>
    <input type="radio" name="usercookies" value="2" <% if request.cookies("aspsky")("usercookies")=2 then response.write "checked" %> >
    1个月
    <input type="radio" name="usercookies" value="3" <% if request.cookies("aspsky")("usercookies")=3 then response.write "checked" %>>
    1年
    <input type="radio" name="usercookies" value="0" <% if request.cookies("aspsky")("usercookies")=0 then response.write "checked" %>>
    不保留 </td>   
      </tr>
<%
dim usersetting,setuserinfo,setusertrue
if rs("usersetting")<>"" then
	usersetting=split(rs("usersetting"),"|||")
	if ubound(usersetting)=1 then
	if isnumeric(usersetting(0)) then setuserinfo=usersetting(0) else setuserinfo=1
	if isnumeric(usersetting(1)) then setusertrue=usersetting(1) else setusertrue=0
	else
	setuserinfo=1
	setusertrue=0
	end if
else
	setuserinfo=1
	setusertrue=0
end if
%>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的基本资料</B>：<BR>开放后别人可以看到您的性别、Email、QQ等信息</td>  <td width=60% class=tablebody1>    
			
              <input type=radio name=setuserinfo value=1  <%if setuserinfo=1 then%>checked<%end if%>>
              开放              <input type=radio name=setuserinfo value=0  <%if setuserinfo=0 then%>checked<%end if%>>
    不开放 </td>   
      </tr>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的真实资料</B>：<BR>开放后别人可以看到您的真实姓名、联系方式等信息</font></td>  <td width=60% class=tablebody1>    
			
              <input type=radio name=setusertrue value=1 <%if setusertrue=1 then%>checked<%end if%>>
              开放              <input type=radio name=setusertrue value=0 <%if setusertrue=0 then%>checked<%end if%>>
    不开放 </td>   
      </tr>
<%
dim userinfo
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,address
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		address=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		address=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	address=""
end if
%>
<tr>
<th height=25 align=left valign=middle colspan=2>&nbsp;个人真实信息（以下内容建议填写）</th>
</tr>
<tr>
<td valign=top width=40% class=tablebody1> 　<b>真实姓名：</b>
<input type=text name=realname size=18 value="<%=realname%>">
</td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>性　格： &nbsp; </b><br>
<% dim KidneyType,theKidney
  KidneyType="多重性格,乐天达观,成熟稳重,幼稚调皮,温柔体帖,活泼可爱,普普通通,内向害羞,外向开朗,心地善良,聪明伶俐,善解人意,风趣幽默,思想开放,积极进取,小心谨慎,郁郁寡欢,正义正直,悲观失意,好吃懒做,处事洒脱,疑神疑鬼,患得患失,异想天开,多愁善感,淡泊名利,见利忘义,瞻前顾后,循规蹈矩,热心助人,快言快语,少言寡语,爱管闲事,追求刺激,豪放不羁,狡猾多变,贪小便宜,见异思迁,情绪多变,水性扬花,重色轻友,胆小怕事,积极负责,勇敢正义,聪明好学,实事求是,务实实际,老实巴交,圆滑老练,脾气暴躁,慢条斯理,诚实坦白,婆婆妈妈,急性子"
  theKidney=split(KidneyType, ",")
	   for i = 0 to ubound(theKidney)	
	   response.write "<input type=""checkbox"" name=""character"" value="""&trim(theKidney(i))&""" "
	 
	       if instr(character,trim(theKidney(i)))>0 then '如果有此项性格
		   response.write "checked" 
		   end if
	  
	   response.write ">"&trim(theKidney(i))&" "
	   if ((i+1) mod 5)=0 then response.write "<br>"  '每行显示六个性格进行换行
	   next   
  %>
</td>
</tr>
<tr>
<td class=tablebody1><b>个人简介： &nbsp;</b><br>
<textarea name=personal rows=6 cols=90% ><%=personal%></textarea>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top width=40%  class=tablebody1>　<b>国　　家：</b>
<b>
<input type=text name=country size=18 value="<%=country%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40%  class=tablebody1>　<b>联系电话：</b>
<b>
<input type=text name=userphone size=18 value="<%=userphone%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>通信地址：</b>
<b>
<input type=text name=address size=18 value="<%=address%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>省　　份：</b>
<input type=text name=province size=18 value="<%=province%>">
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>城　　市：
</b>
<input type=text name=city size=18 value="<%=city%>">
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>生　　肖：
</b>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=鼠 <%if shengxiao="鼠" then%>selected<%end if%>>鼠</option>
<option value=牛 <%if shengxiao="牛" then%>selected<%end if%>>牛</option>
<option value=虎 <%if shengxiao="虎" then%>selected<%end if%>>虎</option>
<option value=兔 <%if shengxiao="兔" then%>selected<%end if%>>兔</option>
<option value=龙 <%if shengxiao="龙" then%>selected<%end if%>>龙</option>
<option value=蛇 <%if shengxiao="蛇" then%>selected<%end if%>>蛇</option>
<option value=马 <%if shengxiao="马" then%>selected<%end if%>>马</option>
<option value=羊 <%if shengxiao="羊" then%>selected<%end if%>>羊</option>
<option value=猴 <%if shengxiao="猴" then%>selected<%end if%>>猴</option>
<option value=鸡 <%if shengxiao="鸡" then%>selected<%end if%>>鸡</option>
<option value=狗 <%if shengxiao="狗" then%>selected<%end if%>>狗</option>
<option value=猪 <%if shengxiao="猪" then%>selected<%end if%>>猪</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>血　　型：</b>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=其他 <%if blood="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>信　　仰：</b>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=佛教 <%if belief="佛教" then%>selected<%end if%>>佛教</option>
<option value=道教 <%if belief="道教" then%>selected<%end if%>>道教</option>
<option value=基督教 <%if belief="基督教" then%>selected<%end if%>>基督教</option>
<option value=天主教 <%if belief="天主教" then%>selected<%end if%>>天主教</option>
<option value=回教 <%if belief="回教" then%>selected<%end if%>>回教</option>
<option value=无神论者 <%if belief="无神论者" then%>selected<%end if%>>无神论者</option>
<option value=共产主义者 <%if belief="共产主义者" then%>selected<%end if%>>共产主义者</option>
<option value=其他 <%if belief="其他" then%>selected<%end if%>>其他</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>职　　业： </b>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="财会/金融" <%if occupation="财会/金融" then%>selected<%end if%>>财会/金融</option>
<option value=工程师 <%if occupation="工程师" then%>selected<%end if%>>工程师</option>
<option value=顾问 <%if occupation="顾问" then%>selected<%end if%>>顾问</option>
<option value=计算机相关行业 <%if occupation="计算机相关行业" then%>selected<%end if%>>计算机相关行业</option>
<option value=家庭主妇 <%if occupation="家庭主妇" then%>selected<%end if%>>家庭主妇</option>
<option value="教育/培训" <%if occupation="教育/培训" then%>selected<%end if%>>教育/培训</option>
<option value="客户服务/支持" <%if occupation="客户服务/支持" then%>selected<%end if%>>客户服务/支持</option>
<option value="零售商/手工工人" <%if occupation="零售商/手工工人" then%>selected<%end if%>>零售商/手工工人</option>
<option value=退休 <%if occupation="退休" then%>selected<%end if%>>退休</option>
<option value=无职业 <%if occupation="无职业" then%>selected<%end if%>>无职业</option>
<option value="销售/市场/广告" <%if occupation="销售/市场/广告" then%>selected<%end if%>>销售/市场/广告</option>
<option value=学生 <%if occupation="学生" then%>selected<%end if%>>学生</option>
<option value=研究和开发 <%if occupation="研究和开发" then%>selected<%end if%>>研究和开发</option>
<option value="一般管理/监督" <%if occupation="一般管理/监督" then%>selected<%end if%>>一般管理/监督</option>
<option value="政府/军队" <%if occupation="政府/军队" then%>selected<%end if%>>政府/军队</option>
<option value="执行官/高级管理" <%if occupation="执行官/高级管理" then%>selected<%end if%>>执行官/高级管理</option>
<option value="制造/生产/操作" <%if occupation="制造/生产/操作" then%>selected<%end if%>>制造/生产/操作</option>
<option value=专业人员 <%if occupation="专业人员" then%>selected<%end if%>>专业人员</option>
<option value="自雇/业主" <%if occupation="自雇/业主" then%>selected<%end if%>>自雇/业主</option>
<option value=其他 <%if occupation="其他" then%>selected<%end if%>>其他</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>婚姻状况：</b>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=未婚 <%if marital="未婚" then%>selected<%end if%>>未婚</option>
<option value=已婚 <%if marital="已婚" then%>selected<%end if%>>已婚</option>
<option value=离异 <%if marital="离异" then%>selected<%end if%>>离异</option>
<option value=丧偶 <%if marital="丧偶" then%>selected<%end if%>>丧偶</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>最高学历：</b>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=小学 <%if education="小学" then%>selected<%end if%>>小学</option>
<option value=初中 <%if education="初中" then%>selected<%end if%>>初中</option>
<option value=高中 <%if education="高中" then%>selected<%end if%>>高中</option>
<option value=大学 <%if education="大学" then%>selected<%end if%>>大学</option>
<option value=硕士 <%if education="硕士" then%>selected<%end if%>>硕士</option>
<option value=博士 <%if education="博士" then%>selected<%end if%>>博士</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>毕业院校：</b>
<input type=text name=college size=18 value="<%=college%>"></td>
</tr>

    <tr align="center"> 
      <td colspan="2" width="100%" class=tablebody2>
            <input type=Submit value="更 新" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="清 除">
      </td>
    </tr>  
    </table></form>
   
</body>
</html> 
<%
end if
rs.close
set rs=nothing
end sub

sub update()
dim CanUseTitle,CanUseTitle1,CanUseTitle2
Dim SplitUserTitle
CanUseTitle=false
CanUseTitle1=false
CanUseTitle2=false
'if request("sex")="" then
'	errmsg=errmsg+"<br>"+"<li>请选择您的性别。"
'	founderr=true
'	exit sub
'elseif request("sex")=0 or request("sex")=1 then
'	sex=request("sex")
'else
'	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
'	founderr=true
'	exit sub
'end if
		
if request("showRe")="" then
	errmsg=errmsg+"<br>"+"<li>请选择有回帖时是否提示您。"
	founderr=true
	exit sub
elseif request("showRe")=0 or request("showRe")=1 then
	showRe=request("showRe")
else
	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
	founderr=true
	exit sub
end if
				
if request("msgFlag")="" then
	errmsg=errmsg+"<br>"+"<li>请选择接收短消息的对象。"
	founderr=true
	exit sub
elseif request("msgFlag")=0 or request("msgFlag")=1 or request("msgFlag")=9 then
	msgFlag=request("msgFlag")
else
	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
	founderr=true
	exit sub
end if

if request("myface")<>"" and ((Cint(forum_setting(54))>0 and myarticle>Cint(forum_setting(54))) or Cint(forum_setting(54))=0) then
	if request("width")="" or request("height")="" then
		errmsg=errmsg+"<br>"+"<li>请输入图片的宽度和高度。"
		founderr=true
		exit sub
	elseif not isInteger(request("width")) or not isInteger(request("height")) then
		errmsg=errmsg+"<br>"+"<li>您输入的字符不合法。"
		founderr=true
		exit sub
	elseif Cint(request("width"))>Cint(forum_setting(57)) then
		errmsg=errmsg+"<br>"+"<li>您输入的图片宽度不符合标准。"
		founderr=true
		exit sub
	elseif Cint(request("height"))>Cint(forum_setting(58)) then
		errmsg=errmsg+"<br>"+"<li>您输入的图片高度不符合标准。"
		founderr=true
		exit sub
	else
		if Cint(forum_setting(55))=0 then
			if instr(lcase(request("myface")),"http://")>0 or instr(lcase(request("myface")),"www.")>0 then
				errmsg=errmsg+"<br>"+"<li>本论坛设置了不允许连接外部网址做头像。"
				founderr=true
				exit sub
			end if
		end if
		face=request("myface")
		width=request("width")
		height=request("height")
	end if
else
	if request("face")="" then
		errmsg=errmsg+"<br>"+"<li>请选择您的个性头像。"
		founderr=true
		exit sub
	else
		face=request("face")
	end if
end if
if width="" or height="" then
	width=forum_setting(38)
	height=forum_setting(39)
end if

if strLength(request("Signature"))>250 then
	errmsg=errmsg+"<br>"+"<li>签名只能在250个字符以内。"
	founderr=true
	exit sub
end if

if request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" then
	birthday=""
else
	birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
	if not isdate(birthday) then birthday=""
end if

dim userinfo,usersetting
userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")

set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
else
	dim faceusemoney,signusemoney
	faceusemoney=0:signusemoney=0
	if trim(lcase(rs("face")))<>trim(lcase(face)) then
		if isnull(myvip) or myvip<>1 then
			faceusemoney=clng(forum_user(22))
		else
			faceusemoney=clng(forum_user(23))
		end if
	end if
	if trim(rs("sign"))<>trim(request("Signature")) then
		if isnull(myvip) or myvip<>1 then
			signusemoney=clng(forum_user(24))
		else
			signusemoney=clng(forum_user(25))
		end if
	end if
	if mymoney<faceusemoney+signusemoney and faceusemoney+signusemoney>0 then
		Errmsg=Errmsg + "<br>" + "<li>您的现金不够修改头像或签名的费用"
		FoundErr=true
		exit sub
	end if
	rs("face")=replace(face,"'","")
	rs("width")=width
	rs("height")=height
	'rs("sex")=sex
	rs("sign")=trim(request("Signature"))
	rs("userwealth")=rs("userwealth")-faceusemoney-signusemoney
	rs("userphoto")=request("userphoto")
	rs("showRe")=showRe
	rs("msgFlag")=msgFlag
	'rs("UserGroup")=request("groupname")
	'判断是否允许提交头衔
	if Cint(Forum_Setting(6))=1 then
		CanUseTitle=true
	end if
	if CanUseTitle and Cint(forum_setting(60))>0 and rs("article")>Cint(forum_setting(60)) then
		CanUseTitle1=true
	elseif CanUseTitle and Cint(forum_setting(60))=0 then
		CanUseTitle1=true
	else
		CanUseTitle1=false
	end if
	if CanUseTitle and Cint(forum_setting(61))>0 and DateDiff("d",rs("AddDate"),Now())>Cint(forum_setting(61)) then
		CanUseTitle2=true
	elseif CanUseTitle and Cint(forum_setting(61))=0 then
		CanUseTitle2=true
	else
		CanUseTitle2=false
	end if
	if CanUseTitle and Cint(forum_setting(62))=1 then
		if CanUseTitle1 and CanUseTitle2 then
			CanUseTitle=true
		else
			CanUseTitle=false
		end if
	elseif CanUseTitle and (CanUseTitle1 or CanUseTitle2) then
		CanUseTitle=true
	else
		CanUseTitle=false
	end if
	if not isnull(request("title")) and CanUseTitle then
		if trim(forum_setting(63))<>"" then
		SplitUserTitle=split(forum_setting(63),"|")
		for i=0 to ubound(SplitUserTitle)
			if instr(lcase(request("title")),lcase(SplitUserTitle(i)))>0 then
				Errmsg=Errmsg + "<br>" + "<li>您设定的头衔中包含了系统禁止使用的字符"
				FoundErr=true
				exit sub
			end if
		next
		end if
		if len(request("title"))>Cint(forum_setting(59)) then
			Errmsg=Errmsg + "<br>" + "<li>您设置的头衔超过了本论坛允许的长度，头衔最大长度为"&forum_setting(59)&"字节。"
			FoundErr=true
			exit sub
		end if
		rs("title")=htmlencode(request("title"))
	end if
	if birthday<>"" then
		rs("birthday")=birthday
	end if
	rs("userinfo")=userinfo
	rs("usersetting")=usersetting
	rs("bbstype")=request("bbstype")
	rs.Update
	usercookies=request("usercookies")
	select case usercookies
	case 0
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 1
		Response.Cookies("aspsky").Expires=Date+1
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 2
		Response.Cookies("aspsky").Expires=Date+31
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 3
		Response.Cookies("aspsky").Expires=Date+365
		Response.Cookies("aspsky")("usercookies") = usercookies
	end select			
		
	Response.Cookies("aspsky")("tempid") = request("bbstype")
	Response.Cookies("aspsky").path=cookiepath
end if
rs.close
set rs=nothing
'******************
'对上传头象进行过滤与改名
if Cint(Forum_Setting(7))=1 then 
	on error resume next
    dim objFSO,upfilename,newfilename
    dim upface,memberid
    set rs=conn.execute("select userid,face from [user] where userid="&userid)
    memberid=rs(0)
    upface=trim(rs(1))
    newfilename=""
	upfilename=split(upface,"/")
	if ubound(upfilename)=1 and upfilename(0)="uploadFace"  then
		if  instr(upfilename(1),"_")=0 then
			newfilename="uploadFace/"&memberid&"_"&upfilename(1)
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			if objFSO.fileExists(Server.MapPath(upface)) then
				objFSO.movefile ""&Server.MapPath(upface)&"",""&Server.MapPath(newfilename)&""
			end if
			If Err.Number = 0 Then
				conn.execute("update [user] set face='"&replace(newfilename,"'","")&"' where userid="&userid)
			end if
			set objFSO=nothing
		end if
	end if
	rs.close
	set rs=nothing
end if
'对上传头象进行过滤与改名结束
'****************
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>