<%@ LANGUAGE=VBScript codepage ="936" %>
<%aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if request("action")="edit" then
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
	rs1.Open "Select * From �ռ��û� where ����='"& aqjh_name &"'", conn1, 1,1
	if rs1.eof and rs1.bof then
		rs1.close
		set rs1=nothing
		conn1.close
		set conn1=nothing
		response.redirect "reg.asp"
		Response.End 
	end if
	mycz="ȷ���޸�"
	rjname=rs1("�ռǱ���")
	pass=rs1("����")
	sm=rs1("˵��")
	lb1=rs1("lb1")
	lb2=rs1("lb2")
	lb3=rs1("lb3")
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
else
	mycz="��ͬ�����Э�鲢ע��"
	rjname=""
	pass=""
	sm=""
	lb1="�����ռ�1"
	lb2="�����ռ�1"
	lb3="�����ռ�1"
end if
%>
<html><head><title>���齭��-�ռǸ���ע��</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=JavaScript>
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit)field.value = field.value.substring(0, maxlimit);
else countfield.value = maxlimit - field.value.length;}
function check()
{var rjname=document.myform.rjname.value;
var sm=document.myform.sm.value;
var pass=document.myform.pass.value;
var lb1=document.myform.lb1.value;
var lb2=document.myform.lb2.value;
var lb3=document.myform.lb3.value;
if(rjname=="" || sm=="" || pass=="" || lb1=="" || lb2=="" || lb3==""){alert("��ʾ�����ݲ���Ϊ�գ�");return false;}
}</script>
</head><LINK href="images/diary.css" rel=stylesheet type=text/css><body bgcolor="#FFFFFF" text="#000000">
<table align=center border=0 cellpadding=0 cellspacing=0 width=574><tbody>
<tr><td><img height=13 src="images/top.gif" width=574></td></tr>
<tr><td align=right background=images/top1.gif><table border=0 cellpadding=0 cellspacing=0 width="94%">
<tbody><tr><td align=right><table border=0 cellpadding=0 cellspacing=0 width="80%"><tbody>
<tr align=right><td><a href="mydiary.asp"><img alt=�����ҵ��ռǱ� border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=�����ռ� border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td><td><a href="top.asp?fs=2"><img alt=�ռ����� border=0 height=21 src="images/anniu4.gif" width=78></a></td>
</tr></tbody></table></td><td rowspan=3 valign=bottom width=26><img height=289 src="images/right.gif" width=8></td>
</tr><tr></tr><tr><td align=middle valign=top>
<form action="diary.asp?action=reg" method=post name=myform onSubmit="return check(this);">
<table bgcolor=#e0d2c5 border=0 cellpadding=3 cellspacing=1 width="88%">
<tr><td align=right bgcolor=#ede5dd width="20%">�ռǱ�����</td>
<td align=left bgcolor=#ede5dd colspan="2">
                    <input name=rjname type=text size="15" maxlength="10"  value="<%=rjname%>">
</td></tr><tbody><tr><td align=right bgcolor=#ede5dd valign="top" width="20%">�ռ����</td>
<td align=left bgcolor=#ede5dd colspan="2"><textarea  cols=40 name=sm rows=5 onKeyDown="textCounter(this.form.sm,this.form.remLen,150);" onKeyUp="textCounter(this.form.sm,this.form.remLen,150);"><%=sm%></textarea></td>
</tr><tr><td align=right bgcolor=#ede5dd height=6 width="20%">�������룺</td><td bgcolor=#ede5dd height=6 width="40%">
<input name=pass type=password size="10" maxlength="10" value="<%=pass%>"></td><td bgcolor=#ede5dd height=6 width="40%">�������룺<input readonly type=text name=remLen size=3 maxlength=3 value="150" style="BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; COLOR: red" >��</td>
</tr><tr align=middle><td bgcolor=#ede5dd width="20%" height=40><div align="right">������ǩ�� </div>
</td><td bgcolor=#ede5dd align="left" height=40 colspan=2>
<input name=lb1 size=12 value="<%=lb1%>" maxlength="12"><input name=lb2 size=12 value="<%=lb2%>" maxlength="12"><input name=lb3 size=12 value="<%=lb2%>" maxlength="12">
</td></tr><tr><td bgcolor=#ede5dd colspan=3 height=40>
<p>�û�������ȫͬ�����·�������:<br>
1. ��վ10��վ����Ȩ�鿴���������ռǣ�վ���������ռǱ������վ���Ȩ��<br>
<br>
2. ��վ�����ֹ�ϴ�������ɫ�顢�Լ������Ϲ��ҹ涨�����£�һ�����֣�����ȡ����ʹ���ʸ�<br>
<br>
3. �ռ�ӵ���߶��Լ����ռ����ݸ��𣬰���һ����������Ϊֱ�ӻ��ӵ��µ����»����·������Ρ�<br>
<br>
4. ��Ϊ����ʧ��������ͨ���������ֶλ�ȡ�����ռ����ݣ�����վ����Ϊ�˳е��κ����Ρ�<br>
<br>
5. ��������ռ�������������δ�������£�����ȡ���ʸ�</p>
<p align="center">
<input name=Submit type=submit value="<%=mycz%>">
</p>
</td>
</tr></tbody></table></form></td></tr></tbody></table></td></tr><tr>
<td><img height=43 src="images/down.gif" width=574></td></tr></tbody></table></body></html>