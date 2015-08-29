<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/z_school_char.asp" -->
<%
'=========================================================
' File: z_school_inclass.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
dim z_membername
if request("username")<>"" then
	z_membername=checkstr(request("username"))
else
	z_membername=membername
end if
if request("action")="out" then
stats="退出同学录"
elseif request("action")="xg" or request("action")="xg1" then
stats="修改资料"
else
stats="加入同学录"
end if
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区同学录的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
call dvbbs_error()
else
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=2><a href=show.asp?filetype=1&boardid="&boardid&">班级相册</a> | <a href=z_school_classuser.asp?boardid="&boardid&">班级成员</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>群体短信</a> | <a href=list.asp?boardid="&boardid&">班级论坛</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>班级主页</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">我的资料</a> | <a href=announce.asp?boardid="&boardid&">我要留言</a> | <a href=z_school_inclass.asp?boardid="&boardid&">加入班级</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">退出班级</a> | <a href=z_school_admin.asp?boardid="&boardid&">班级管理</a></tD></tr>"

if request("action")="out" then
response.write outschool(boardid,z_membername)
elseif request("action")="updat" or request("action")="xg1" then
call update()
else
call inclasszl()
end if


if founderr then call dvbbs_error()
response.write "</TABLE>"
end if
call activeonline()
call footer()



Rem 用户加入同学录
function inschool(boardid,username)
dim boarduser,inschool1
inschool1=true
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	inschool="错误版面参数！"
	inschool1=false
elseif cint(right(rs(1),1))=0 then
	inschool="错误同学录版面参数！"
	inschool1=false
elseif isnull(rs(0)) or rs(0)="" then
	boarduser=username
else
	boarduser=split(rs(0), ",")
	for i = 0 to ubound(boarduser)
		if trim(boarduser(i))=trim(username) then
			inschool="错误！请不要重复加入该同学录！"
			inschool1=false
		exit for
		end if
	next
	boarduser=username&","&rs(0)
end if
if inschool1 then
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	inschool="你已成功加入该同学录！"
end if
rs.close
set rs=nothing
inschool="<TR><TH height=25 colspan=2>系 统 提 示</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>"&inschool&"</td></tr>"
end function


Rem 用户退出同学录
function outschool(boardid,username)
dim boarduser,inschool1
inschool1=0
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	outschool="错误版面参数！"
	inschool1=1
elseif cint(right(rs(1),1))=0 then
	outschool="错误同学录版面参数！"
	inschool1=1
elseif isnull(rs(0)) or rs(0)="" then
	outschool="错误同学录参数！"
	inschool1=1
else
	boarduser=rs(0)
end if
if instr(username,",") <> 0 and inschool1=0 then
inschool1=2
end if
if inschool1=0 then
	username=" "&username&" "
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	boarduser=replace(boarduser,username," ")
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	outschool="你已成功退出该同学录！"
elseif inschool1=2 then
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	dim dede
	username=split(username, ",")
	for each dede in username
		dede=" "&dede&" "
		boarduser=replace(boarduser,dede," ")
	next
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	outschool="你已成功退出该同学录！"
end if
outschool="<TR><TH height=25 colspan=2>系 统 提 示</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>"&outschool&"</td></tr>"
end function

sub inclasszl()
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
end if%>
<TR><TH height=25 colspan=2>填 写 资 料</TH></tR>
<form action="z_school_inclass.asp?action=<%if request("action")="xg" then%>xg1<%else%>updat<%end if%>&boardid=<%=boardid%>" method=POST name="theForm">
<%dim usersetting,setuserinfo,setusertrue
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
end if%>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的基本资料</B>：<BR>开放后别人可以看到您的性别、Email、QQ等信息</td>
<td width=60% class=tablebody1>&nbsp;<input type=radio name=setuserinfo value=1  <%if setuserinfo=1 then%>checked<%end if%>>&nbsp;开放&nbsp;&nbsp;<input type=radio name=setuserinfo value=0  <%if setuserinfo=0 then%>checked<%end if%>>&nbsp;不开放</td>
</tr>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的真实资料</B>：<BR>开放后别人可以看到您的真实姓名、联系方式等信息</font></td>
<td width=60% class=tablebody1>&nbsp;<input type=radio name=setusertrue value=1 <%if setusertrue=1 then%>checked<%end if%>>&nbsp;开放&nbsp;&nbsp;<input type=radio name=setusertrue value=0 <%if setusertrue=0 then%>checked<%end if%>>&nbsp;不开放</td>   
</tr>
<%dim userinfo
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
end if%>
<tr>
<th height=25 align=left valign=middle colspan=2>&nbsp;个人真实信息（以下内容建议填写）</th>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>真实姓名：</b><input type=text name=realname size=18 value="<%=realname%>">&nbsp;<font color=red>(必填)</font></td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>性　格： &nbsp; </b><br><%
	dim KidneyType,theKidney
	KidneyType="多重性格,乐天达观,成熟稳重,幼稚调皮,温柔体贴,活泼可爱,普普通通,内向害羞,外向开朗,心地善良,聪明伶俐,善解人意,风趣幽默,思想开放,积极进取,小心谨慎,郁郁寡欢,正义正直,悲观失意,好吃懒做,处事洒脱,疑神疑鬼,患得患失,异想天开,多愁善感,淡泊名利,见利忘义,瞻前顾后,循规蹈矩,热心助人,快言快语,少言寡语,爱管闲事,追求刺激,豪放不羁,狡猾多变,贪小便宜,见异思迁,情绪多变,水性扬花,重色轻友,胆小怕事,积极负责,勇敢正义,聪明好学,实事求是,务实实际,老实巴交,圆滑老练,脾气暴躁,慢条斯理,诚实坦白,婆婆妈妈,急性子"
	theKidney=split(KidneyType, ",")
	for i = 0 to ubound(theKidney)	
		response.write "<input type=""checkbox"" name=""character"" value="""&trim(theKidney(i))&""" "
		if instr(character,trim(theKidney(i)))>0 then '如果有此项性格
			response.write "checked" 
		end if
		response.write ">"&trim(theKidney(i))&" "
		if ((i+1) mod 5)=0 then response.write "<br>"  '每行显示六个性格进行换行
	next   
%></td>
</tr>
<tr>
<td class=tablebody1><b>个人简介： &nbsp;</b><br><textarea name=personal rows=5 cols=90% ><%=personal%></textarea></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>国　　家：</b><b><input type=text name=country size=18 value="<%=country%>"></b></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>联系电话：</b><b><input type=text name=userphone size=18 value="<%=userphone%>"></b>&nbsp;<font color=red>(必填)</font></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>通信地址：</b><b><input type=text name=address size=18 value="<%=address%>"></b>&nbsp;<font color=red>(必填)</font></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>省　　份：</b><input type=text name=province size=18 value="<%=province%>"></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>城　　市：</b><input type=text name=city size=18 value="<%=city%>"></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>生　　肖：</b><select size=1 name=shengxiao>
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
<td valign=top width=40% class=tablebody1>&nbsp;<b>血　　型：</b><select size=1 name=blood>
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
<td valign=top width=40% class=tablebody1>&nbsp;<b>信　　仰：</b><select size=1 name=belief>
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
<td valign=top width=40% class=tablebody1>&nbsp;<b>职　　业：</b><select name=occupation>
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
<td valign=top width=40% class=tablebody1>&nbsp;<b>婚姻状况：</b><select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=未婚 <%if marital="未婚" then%>selected<%end if%>>未婚</option>
<option value=已婚 <%if marital="已婚" then%>selected<%end if%>>已婚</option>
<option value=离异 <%if marital="离异" then%>selected<%end if%>>离异</option>
<option value=丧偶 <%if marital="丧偶" then%>selected<%end if%>>丧偶</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>最高学历：</b><select size=1 name=education>
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
<td valign=top width=40% class=tablebody1>&nbsp;<b>毕业院校：</b><input type=text name=college size=18 value="<%=college%>"></td>
</tr>
<tr height=25>
<td valign=top width=40% class=tablebody1>&nbsp;<B>OICQ号码：</B><input type=TEXT name="OICQ" size=18 value="<%if rs("oicq")<>"" then%><%=rs("oicq")%><%end if%>" maxlength=20></td>   
</tr>
<tr align="center"> 
<td colspan="2" width="100%" class=tablebody2><input type=Submit value="加 入" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="清 除"></td>
</tr>
</form>
<%end sub


sub update()
if request("address")="" then
	errmsg=errmsg+"<br>"+"<li>通信地址在同学录是必填的。"
	founderr=true
	exit sub
end if
if trim(request("oicq"))<>"" then
	if not isnumeric(request("oicq")) or len(request("oicq"))>12 then
		errmsg=errmsg+"<br>"+"<li>Oicq号码只能是4-12位数字，在同学录是必填的。"
		founderr=true
		exit sub
	end if
end if
dim realname
realname=trim(request("realname"))
realname=replace(realname," ","")
if realname="" or strLength(realname)<4 or strLength(realname)>8 then
	errmsg=errmsg+"<br>"+"<li>真实姓名输入有误。"
	founderr=true
exit sub
elseif not isChinese(realname) then
	errmsg=errmsg+"<br>"+"<li>真实姓名应为汉字。"
	founderr=true
	exit sub
end if
if request.form("userphone")="" then
	errmsg=errmsg+"<br>"+"<li>联系电话在同学录是必填的。"
	founderr=true
	exit sub
end if
dim userinfo,usersetting
userinfo=checkreal(realname) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")
if not isnull(request.form("oicq")) and request.form("oicq")<>"" then
	conn.execute("update [User] set oicq='"&request.form("oicq")&"' where userid="&userid)
end if
conn.execute("update [User] set userinfo='"&userinfo&"' where userid="&userid)
conn.execute("update [User] set usersetting='"&usersetting&"' where userid="&userid)
if request("action")="xg1" then
response.write "<TR><TH height=25 colspan=2>系 统 提 示</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>修改成功！</td></tr>" 
else
response.write inschool(boardid,z_membername)
end if
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>
