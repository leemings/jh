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
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ڸ�ʲô��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From �ռ��û� where ����='"& aqjh_name &"'", conn1, 1,1
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
mycz="�ύ"
bmcs="�����ռ�"
face=5
weather=1
n=Year(date())
y=Month(date())
r=Day(date())
mm=1
lb=1
if id<>"" then
	rs1.close
	rs1.open "select * from �ռ� where id="&id,conn1,1,1
	name=rs1("�û���")
	if name<>aqjh_name then
		rs1.close
		set rs1=nothing
		conn1.close
		set conn1=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ������Ȩ�༭�ռǣ�');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	mycz="�༭"
	title=rs1("�ռǱ���")
	diary=rs1("�ռ�����")
	weather=rs1("����")
	face=rs1("����")
	mm=rs1("����")
	bmcs=rs1("��������")
	rq=rs1("��д����")
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
<html><head><title>�ռǱ༭</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css>
<script src="ubbcode.js"></script>
<SCRIPT language=JavaScript>
var mymm=1;
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit)field.value = field.value.substring(0, maxlimit);
else countfield.value = maxlimit - field.value.length;}
function help(){alert("����:\n����:�κ��˶����Թۿ�.\n\n��ȫ���ܣ�˭Ҳ������\n\n����:ֻ��֪������ſ���.\n\nָ����:ֻ����(��)�ſ���.\n\nָ���ȼ�:�ȼ��������.\n\nָ������:������ɱ���\n");}
function a1(){mymm=1;document.myform.bmcs.value="�����ռ�";}
function a2(){mymm=2;document.myform.bmcs.value="˭Ҳ���ÿ�";}
function a3(){mymm=3;document.myform.bmcs.value="����������";}
function a4(){mymm=4;document.myform.bmcs.value="��������";}
function a5(){mymm=5;document.myform.bmcs.value="���Ƶȼ�";}
function a6(){mymm=6;document.myform.bmcs.value="��������";}
function check()
{var bmcs=document.myform.bmcs.value;var diary=document.myform.diary.value;var title=document.myform.title.value;
var diary=document.myform.diary.value;
var bmcs=document.myform.bmcs.value;
var pattern = /^([0-9])+$/;
if(mymm=="5"){if(pattern.test(bmcs)!=true){alert("��ʾ�����Ƶȼ�ʱ,�������֣�");return false;}}
if(bmcs=="" || diary=="" || title==""){alert("��ʾ�����ݲ���Ϊ�գ�");return false;}
if(title.length>50){alert("��ʾ����������̫�࣡");return false;}
if(bmcs.length>10){alert("��ʾ�����ܲ���̫�࣡");return false;}
if(title.length<5){alert("��ʾ����������̫�٣�");return false;}
if(diary.length<50){alert("��ʾ���ռ�����̫�٣�");return false;}}
</script>
<body bgcolor="#FFFFFF" text="#000000"><table align=center border=0 cellpadding=0 cellspacing=0 idth=574>
<tbody><tr><td><img height=13 src="images/top.gif" width=574></td>
</tr><tr><td align=right background=images/top1.gif><table border=0 cellpadding=0 cellspacing=0 width="94%">
<tbody><tr><td align=right>
<table border=0 cellpadding=0 cellspacing=0 width="80%"><tbody>
<tr align=right><td><a href="mydiary.asp"><img alt=�����ҵ��ռǱ� border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=�����ռ� border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
<td><a href="top.asp?fs=2"><img alt=�ռ����� border=0 height=21 src="images/anniu4.gif" width=78></a></td>
</tr></tbody></table></td>
<td rowspan=3 valign=bottom width=26><img height=289 src="images/right.gif" width=8></td></tr>
<tr><td>&nbsp;</td></tr><tr><td valign=top>
<form action="diary.asp?action=edit&?id=<%=id%>" method=post name=myform onSubmit="return check(this);">
<table bgcolor=#e0d2c5 border=0 cellpadding=2 cellspacing=1 width="98%">
<tbody><tr bgcolor=#ede5dd><td align=right width=71>����</td> <td colspan=2 width=422> <font size="-1">
<input name="mm" type=radio value="1" onclick="a1()" <%if mm=1 then%>checked<%end if%>>����<input name="mm" type=radio value="2" onclick="a2()" <%if mm=2 then%>checked<%end if%>>����<input name="mm" type=radio value="3" onclick="a3()" <%if mm=3 then%>checked<%end if%>>����<input name="mm" type=radio value="4" onclick="a4()" <%if mm=4 then%>checked<%end if%>>ָ����<input name="mm" type=radio value="5" onclick="a5()" <%if mm=5 then%>checked<%end if%>>ָ���ȼ�<input name="mm" type=radio value="6" onclick="a6()" <%if mm=6 then%>checked<%end if%>>ָ������
<a href="#" onclick="help()"><font size="-2" color=blue>��ϸ˵��</font></a><br>���ܲ���:<input name="bmcs" style="readonly=true;" type="text" maxlength="10" size="10" value="<%=bmcs%>"></font></td>
</tr><tr bgcolor=#ede5dd><td align=right width=71>���⣺</td><td colspan=2 width=422><input name=title maxlength="40" size="40" value="<%=title%>">
</td></tr><tr bgcolor=#ede5dd><td align=right width=71>���ڣ�</td><td colspan=2 width=422>
<input maxlength=4 name=my_year size=4 value="<%=n%>">��
<select class=select name=my_month size=1>
<option value=1 <%if y=1 then%>selected<%end if%>>1</option><option value=2 <%if y=2 then%>selected<%end if%>>2</option><option value=3 <%if y=3 then%>selected<%end if%>>3</option><option value=4 <%if y=4 then%>selected<%end if%>>4</option>
<option value=5 <%if y=5 then%>selected<%end if%>>5</option><option value=6 <%if y=6 then%>selected<%end if%>>6</option><option value=7 <%if y=7 then%>selected<%end if%>>7</option><option value=8 <%if y=8 then%>selected<%end if%>>8</option>
<option value=9 <%if y=9 then%>selected<%end if%>>9</option><option value=10 <%if y=10 then%>selected<%end if%>>10</option><option  value=11 <%if y=11 then%>selected<%end if%>>11</option><option value=12 <%if y=12 then%>selected<%end if%>>12</option>
</select>��<select class=select name=my_day size=1><option value=1 <%if r=1 then%>selected<%end if%>>1</option>
<option value=2 <%if r=2 then%>selected<%end if%>>2</option><option value=3 <%if r=3 then%>selected<%end if%>>3</option><option value=4 <%if r=4 then%>selected<%end if%>>4</option><option value=5 <%if r=5 then%>selected<%end if%>>5</option><option value=6 <%if r=6 then%>selected<%end if%>>6</option>
<option value=7 <%if r=7 then%>selected<%end if%>>7</option><option value=8 <%if r=8 then%>selected<%end if%>>8</option><option value=9 <%if r=9 then%>selected<%end if%>>9</option><option value=10 <%if r=10 then%>selected<%end if%>>10</option><option value=11 <%if r=11 then%>selected<%end if%>>11</option>
<option value=12 <%if r=12 then%>selected<%end if%>>12</option><option value=13 <%if r=13 then%>selected<%end if%>>13</option><option value=14 <%if r=14 then%>selected<%end if%>>14</option><option value=15 <%if r=15 then%>selected<%end if%>>15</option><option value=16 <%if r=16 then%>selected<%end if%>>16</option>
<option value=17 <%if r=17 then%>selected<%end if%>>17</option><option value=18 <%if r=18 then%>selected<%end if%>>18</option><option value=19 <%if r=19 then%>selected<%end if%>>19</option><option value=20 <%if r=20 then%>selected<%end if%>>20</option><option value=21 <%if r=21 then%>selected<%end if%>>21</option>
<option value=22 <%if r=22 then%>selected<%end if%>>22</option><option value=23 <%if r=23 then%>selected<%end if%>>23</option><option value=24 <%if r=24 then%>selected<%end if%>>24</option><option value=25 <%if r=25 then%>selected<%end if%>>25</option><option value=26 <%if r=26 then%>selected<%end if%>>26</option>
<option value=27 <%if r=27 then%>selected<%end if%>>27</option><option value=28 <%if r=28 then%>selected<%end if%>>28</option><option value=29 <%if r=29 then%>selected<%end if%>>29</option><option value=30 <%if r=30 then%>selected<%end if%>>30</option><option value=31 <%if r=31 then%>selected<%end if%>>31</option></select>��</td></tr>
<tr bgcolor=#ede5dd><td align=right width=71>������</td><td  width=422>
<select name=weather  onChange="document.images['face'].src='images/wether'+options[selectedIndex].value+'.gif';">
<option value="1" <%if weather=1 then%>selected<%end if%>>����</option><option value="2" <%if weather=2 then%>selected<%end if%>>����</option><option value="3" <%if weather=3 then%>selected<%end if%>>Ʈѩ</option><option value="4" <%if weather=4 then%>selected<%end if%>>����</option></select>
<img id="face" src="images/wether<%=weather%>.gif" width="18" height="18" alt="����">
����ǩ:<input name=lb type=radio value=1 <%if lb=1 then%>checked<%end if%>><%=lb1%><input name=lb type=radio value=2 <%if lb=2 then%>checked<%end if%>><%=lb2%><input name=lb type=radio value=3 <%if lb=3 then%>checked<%end if%>><%=lb3%>
</td></tr><tr bgcolor=#ede5dd><td align=right height=44 width=71>���飺</td><td colspan=2 height=44 width=422>
<input name=face type=radio value=1 <%if face=1 then%>checked<%end if%>>��ŭ<img height=19 src="images/1.gif" width=19>
<input name=face type=radio value=2 <%if face=2 then%>checked<%end if%>>����<img height=19 src="images/2.gif" width=19>
<input name=face type=radio value=3 <%if face=3 then%>checked<%end if%>>����<img height=19 src="images/3.gif" width=19>
<input name=face type=radio value=4 <%if face=4 then%>checked<%end if%>>�Ծ�<img height=19 src="images/4.gif" width=19><br>
<input name=face type=radio value=5 <%if face=5 then%>checked<%end if%>>����<img height=19 src="images/5.gif" width=19>
<input name=face type=radio value=6 <%if face=6 then%>checked<%end if%>>�ѹ�<img height=19 src="images/6.gif" width=19>
<input name=face type=radio value=7 <%if face=7 then%>checked<%end if%>>��Ц<img height=19 src="images/7.gif" width=19>
<input name=face type=radio value=8 <%if face=8 then%>checked<%end if%>>��Ƥ<img height=19 src="images/8.gif" width=19></td>
</tr><tr bgcolor=#ede5dd><td width=-1>UBB����:</td><td colspan=2 width=422><!--#include file="getubb.asp"--> </td>
</tr><tr bgcolor=#ede5dd><td align=right valign=top width=71>����:</td>
<td colspan=2 width=422 bgcolor="#ede5dd"><textarea class=text2 name=diary cols="60" rows="6"  onKeyDown="textCounter(this.form.diary,this.form.remLen,2048);" onKeyUp="textCounter(this.form.diary,this.form.remLen,2048);">
<%=diary%>
</textarea>
</td></tr><tr bgcolor=#ede5dd><td width=71>&nbsp;</td><td align=left height=40 width=422><div align="center">
��������2048�֣��������������룺<input readonly type=text name=remLen size=4 maxlength=3 value="2048" style="BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; COLOR: red" >�ֽ�
<br>
<%if id<>"" then %>
<input  name=id type=hidden value=<%=id%>>
<%end if%>
<input name=Submit type=submit value=<%=mycz%>>&nbsp;&nbsp;<input name=Submit2 type=reset value=����>
</div></td></tr></tbody></table></form></td></tr></tbody></table></td></tr><tr><td><img height=43 src="images/down.gif" width=574></td></tr></tbody></table>
</body></html>