<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")
if (not isnumeric(id)) then 
	Response.Write "<script language=JavaScript>{alert('提示：你在搞什么！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记用户 where 姓名='"& aqjh_name &"'", conn1, 1,1
if rs1.eof and rs1.bof then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Redirect "reg.asp"
end if
lb1=rs1("lb1")
lb2=rs1("lb2")
lb3=rs1("lb3")
mycz="提交"
bmcs="公开日记"
face=5
weather=1
n=Year(date())
y=Month(date())
r=Day(date())
mm=1
lb=1
if id<>"" then
	rs1.close
	rs1.open "select * from 日记 where id="&id,conn1,1,1
	name=rs1("用户名")
	if name<>aqjh_name and aqjh_grade<>10 then
		rs1.close
		set rs1=nothing
		conn1.close
		set conn1=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你无权编辑日记！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	mycz="编辑"
	title=rs1("日记标题")
	diary=rs1("日记内容")
	weather=rs1("天气")
	face=rs1("心情")
	mm=rs1("保密")
	bmcs=rs1("保密条件")
	rq=rs1("书写日期")
	n=Year(rq)
	y=Month(rq)
	r=Day(rq)
	lb=rs1("lb")
end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html><head><title>日记书写</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<script src="ubbcode.js"></script>
<SCRIPT language=JavaScript>
var mymm=1;
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit)field.value = field.value.substring(0, maxlimit);
else countfield.value = maxlimit - field.value.length;}
function help(){alert("帮助:\n公开:任何人都可以观看.\n\n完全保密：谁也看不到\n\n密码:只有知道密码才可以.\n\n指定人:只有他(她)才可以.\n\n指定等级:等级不够你别看.\n\n指定门派:别的门派别看了\n");}
function a1(){mymm=1;document.myform.bmcs.value="公开日记";}
function a2(){mymm=2;document.myform.bmcs.value="谁也不让看";}
function a3(){mymm=3;document.myform.bmcs.value="请输入密码";}
function a4(){mymm=4;document.myform.bmcs.value="输入人名";}
function a5(){mymm=5;document.myform.bmcs.value="限制等级";}
function a6(){mymm=6;document.myform.bmcs.value="门派名字";}
function check()
{var bmcs=document.myform.bmcs.value;var diary=document.myform.diary.value;var title=document.myform.title.value;
var diary=document.myform.diary.value;
var bmcs=document.myform.bmcs.value;
var pattern = /^([0-9])+$/;
if(mymm=="5"){if(pattern.test(bmcs)!=true){alert("提示：限制等级时,请用数字！");return false;}}
if(bmcs=="" || diary=="" || title==""){alert("提示：数据不能为空！");return false;}
if(title.length>50){alert("提示：标题文字太多！");return false;}
if(bmcs.length>10){alert("提示：保密参数太多！");return false;}
if(title.length<5){alert("提示：标题文字太少！");return false;}
if(diary.length<50){alert("提示：日记内容太少！");return false;}}
</script>
<body bgcolor="#FFFFFF" text="#000000"><table align=center border=0 cellpadding=0 cellspacing=0 idth=574>
<tbody><tr><td><img height=13 src="images/top.gif" width=574></td>
</tr><tr><td align=right background=images/top1.gif><table border=0 cellpadding=0 cellspacing=0 width="94%">
<tbody><tr><td align=right>
<table border=0 cellpadding=0 cellspacing=0 width="80%"><tbody>
<tr align=right><td><a href="mydiary.asp"><img alt=进入我的日记本 border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=最新日记 border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
<td><a href="top.asp?fs=2"><img alt=日记排行 border=0 height=21 src="images/anniu4.gif" width=78></a></td>
</tr></tbody></table></td>
<td rowspan=3 valign=bottom width=26><img height=289 src="images/right.gif" width=8></td></tr>
<tr><td>&nbsp;</td></tr><tr><td valign=top>
<form action="diary.asp?action=add" method=post name=myform onSubmit="return check(this);">
<table bgcolor=#e0d2c5 border=0 cellpadding=2 cellspacing=1 width="98%">
<tbody><tr bgcolor=#ede5dd><td align=right width=71>保密</td> <td colspan=2 width=422> <font size="-1">
<input name="mm" type=radio value="1" onclick="a1()" <%if mm=1 then%>checked<%end if%>>公开<input name="mm" type=radio value="2" onclick="a2()" <%if mm=2 then%>checked<%end if%>>保密<input name="mm" type=radio value="3" onclick="a3()" <%if mm=3 then%>checked<%end if%>>密码<input name="mm" type=radio value="4" onclick="a4()" <%if mm=4 then%>checked<%end if%>>指定人<input name="mm" type=radio value="5" onclick="a5()" <%if mm=5 then%>checked<%end if%>>指定等级<input name="mm" type=radio value="6" onclick="a6()" <%if mm=6 then%>checked<%end if%>>指定门派
<a href="#" onclick="help()"><font size="-2" color=blue>详细说明</font></a><br>保密参数:<input name="bmcs" style="readonly=true;" type="text" maxlength="10" size="10" value="<%=bmcs%>"></font></td>
</tr><tr bgcolor=#ede5dd><td align=right width=71>标题：</td><td colspan=2 width=422><input name=title maxlength="40" size="40" value="<%=title%>">
</td></tr><tr bgcolor=#ede5dd><td align=right width=71>日期：</td><td colspan=2 width=422>
<input maxlength=4 name=my_year size=4 value="<%=n%>">年
<select class=select name=my_month size=1>
<option value=1 <%if y=1 then%>selected<%end if%>>1</option><option value=2 <%if y=2 then%>selected<%end if%>>2</option><option value=3 <%if y=3 then%>selected<%end if%>>3</option><option value=4 <%if y=4 then%>selected<%end if%>>4</option>
<option value=5 <%if y=5 then%>selected<%end if%>>5</option><option value=6 <%if y=6 then%>selected<%end if%>>6</option><option value=7 <%if y=7 then%>selected<%end if%>>7</option><option value=8 <%if y=8 then%>selected<%end if%>>8</option>
<option value=9 <%if y=9 then%>selected<%end if%>>9</option><option value=10 <%if y=10 then%>selected<%end if%>>10</option><option  value=11 <%if y=11 then%>selected<%end if%>>11</option><option value=12 <%if y=12 then%>selected<%end if%>>12</option>
</select>月<select class=select name=my_day size=1><option value=1 <%if r=1 then%>selected<%end if%>>1</option>
<option value=2 <%if r=2 then%>selected<%end if%>>2</option><option value=3 <%if r=3 then%>selected<%end if%>>3</option><option value=4 <%if r=4 then%>selected<%end if%>>4</option><option value=5 <%if r=5 then%>selected<%end if%>>5</option><option value=6 <%if r=6 then%>selected<%end if%>>6</option>
<option value=7 <%if r=7 then%>selected<%end if%>>7</option><option value=8 <%if r=8 then%>selected<%end if%>>8</option><option value=9 <%if r=9 then%>selected<%end if%>>9</option><option value=10 <%if r=10 then%>selected<%end if%>>10</option><option value=11 <%if r=11 then%>selected<%end if%>>11</option>
<option value=12 <%if r=12 then%>selected<%end if%>>12</option><option value=13 <%if r=13 then%>selected<%end if%>>13</option><option value=14 <%if r=14 then%>selected<%end if%>>14</option><option value=15 <%if r=15 then%>selected<%end if%>>15</option><option value=16 <%if r=16 then%>selected<%end if%>>16</option>
<option value=17 <%if r=17 then%>selected<%end if%>>17</option><option value=18 <%if r=18 then%>selected<%end if%>>18</option><option value=19 <%if r=19 then%>selected<%end if%>>19</option><option value=20 <%if r=20 then%>selected<%end if%>>20</option><option value=21 <%if r=21 then%>selected<%end if%>>21</option>
<option value=22 <%if r=22 then%>selected<%end if%>>22</option><option value=23 <%if r=23 then%>selected<%end if%>>23</option><option value=24 <%if r=24 then%>selected<%end if%>>24</option><option value=25 <%if r=25 then%>selected<%end if%>>25</option><option value=26 <%if r=26 then%>selected<%end if%>>26</option>
<option value=27 <%if r=27 then%>selected<%end if%>>27</option><option value=28 <%if r=28 then%>selected<%end if%>>28</option><option value=29 <%if r=29 then%>selected<%end if%>>29</option><option value=30 <%if r=30 then%>selected<%end if%>>30</option><option value=31 <%if r=31 then%>selected<%end if%>>31</option></select>日</td></tr>
<tr bgcolor=#ede5dd><td align=right width=71>天气：</td><td  width=422>
<select name=weather  onChange="document.images['face'].src='images/wether'+options[selectedIndex].value+'.gif';">
<option value="1" <%if weather=1 then%>selected<%end if%>>晴天</option><option value="2" <%if weather=2 then%>selected<%end if%>>下雨</option><option value="3" <%if weather=3 then%>selected<%end if%>>飘雪</option><option value="4" <%if weather=4 then%>selected<%end if%>>多云</option></select>
<img id="face" src="images/wether<%=weather%>.gif" width="18" height="18" alt="天气">
▲书签:<input name=lb type=radio value=1 <%if lb=1 then%>checked<%end if%>><%=lb1%><input name=lb type=radio value=2 <%if lb=2 then%>checked<%end if%>><%=lb2%><input name=lb type=radio value=3 <%if lb=3 then%>checked<%end if%>><%=lb3%>
</td></tr><tr bgcolor=#ede5dd><td align=right height=44 width=71>心情：</td><td colspan=2 height=44 width=422>
<input name=face type=radio value=1 <%if face=1 then%>checked<%end if%>>愤怒<img height=19 src="images/1.gif" width=19>
<input name=face type=radio value=2 <%if face=2 then%>checked<%end if%>>困惑<img height=19 src="images/2.gif" width=19>
<input name=face type=radio value=3 <%if face=3 then%>checked<%end if%>>害羞<img height=19 src="images/3.gif" width=19>
<input name=face type=radio value=4 <%if face=4 then%>checked<%end if%>>吃惊<img height=19 src="images/4.gif" width=19><br>
<input name=face type=radio value=5 <%if face=5 then%>checked<%end if%>>开心<img height=19 src="images/5.gif" width=19>
<input name=face type=radio value=6 <%if face=6 then%>checked<%end if%>>难过<img height=19 src="images/6.gif" width=19>
<input name=face type=radio value=7 <%if face=7 then%>checked<%end if%>>大笑<img height=19 src="images/7.gif" width=19>
<input name=face type=radio value=8 <%if face=8 then%>checked<%end if%>>调皮<img height=19 src="images/8.gif" width=19></td>
</tr><tr bgcolor=#ede5dd><td width=-1>UBB代码:</td><td colspan=2 width=422><!--#include file="getubb.asp"--> </td>
</tr><tr bgcolor=#ede5dd><td align=right valign=top width=71>正文:</td>
<td colspan=2 width=422 bgcolor="#ede5dd"><textarea class=text2 name=diary cols="60" rows="6"  onKeyDown="textCounter(this.form.diary,this.form.remLen,2048);" onKeyUp="textCounter(this.form.diary,this.form.remLen,2048);">
<%=diary%>
</textarea>
</td></tr><tr bgcolor=#ede5dd><td width=71>&nbsp;</td><td align=left height=40 width=422><div align="center">
限制输入2048字，现在您还可输入：<input readonly type=text name=remLen size=4 maxlength=3 value="2048" style="BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; COLOR: red" >字节
<br>
<%if id<>"" then %>
<input  name=id type=hidden value=<%=id%>>
<%end if%>
<input name=Submit type=submit value=<%=mycz%>>&nbsp;&nbsp;<input name=Submit2 type=reset value=撤消>
</div></td></tr></tbody></table></form></td></tr></tbody></table></td></tr><tr><td><img height=43 src="images/down.gif" width=574></td></tr></tbody></table>
</body></html>