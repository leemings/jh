<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_Visual_const.asp"-->
<%
dim shopid,curpage,sex,rsvisual,sqlvisual,CurType,CurSubType
dim isHas
dim myLayerStr,myvisualsplit
dim PhotoPrice1,PhotoPrice2,PhotoPeriod,PhotoStatus,PhotoUserCount,PriceMin,AccoutermentType

if not founduser then
	errmsg=errmsg+"<br>"+"<li>����û��<a href=login.asp>��¼��̳</a>������ʹ�ø���������ƹ��ܡ��������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	founderr=true
end if

myvisualsplit=split(v_myvisual,"|")
myLayerStr=GetLayerStr(myvisualsplit)

shopid=checkstr(request("shopid"))
if not isnumeric(shopid) or isnull(shopid) or shopid="" then 
	shopid=198
elseif shopid<1 then 
	shopid=198
end if
CurType=shopid mod 100
CurSubType=shopid \ 100
sex=checkstr(request("sex"))
if isnull(sex) then sex=""
curPage=request("page")
if curpage="" or not isInteger(curpage) then
	curpage=1
else
	curpage=clng(curpage)
end if
stats="�����������"
call nav()
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
	%><!--#include file="z_controlpanel.asp"--><br><%
	if founderr then call dvbbs_error()
	call main()
end if
call activeonline()
call footer()

sub main()
	dim minsubtype,totalrec,pagecount,curcount
	dim TypeCount,TypeName,SubTypeName,arraystr,lastseq
	dim SearchMsg
	set rsvisual=server.createobject("ADODB.Recordset")
	sqlvisual="select top 1 photoprice1,photoprice2,photoperiod,photostatus,photousercount,PriceMin,Accouterment from visual_config"
	rsvisual.open sqlvisual,conn,1,1
	if not (rsvisual.eof or rsvisual.bof) then
		PhotoPrice1=rsvisual(0)
		PhotoPrice2=rsvisual(1)
		PhotoPeriod=rsvisual(2)
		PhotoStatus=rsvisual(3)
		PhotoUserCount=rsvisual(4)
		PriceMin=rsvisual(5)
		AccoutermentType=rsvisual(6)
	else
		PhotoPrice1=20
		PhotoPrice2=10
		PhotoPeriod=5
		PhotoStatus=1
		PhotoUserCount=5
		PriceMin=5
	end if
	rsvisual.close
	TypeName=""
	SubTypeName=""
	response.write "<table cellpadding=2 cellspacing=1 align=center class=tableborder1>"
	response.write "<tr height=25>"
	response.write "<th valign=middle id=tabletitlelink>"
	if CurType=0 then
		response.write "<font color="&forum_body(8)&">"
		TypeName="�ҵĴ����"
	else
		response.write "<a href=?shopid=100&sex="&sex&"&page=1>"
	end if
	response.write "�ҵĴ����"
	if CurType=0 then
		response.write "</font>"
	else
		response.write "</a>"
	end if
	response.write "</th>"
	response.write "<th valign=middle id=tabletitlelink>"
	if CurType=99 then
		response.write "<font color="&forum_body(8)&">"
		TypeName="�Ƽ���Ʒ"
	else
		response.write "<a href=?shopid=199&sex="&sex&"&page=1>"
	end if
	response.write "�Ƽ���Ʒ"
	if CurType=99 then
		response.write "</font>"
	else
		response.write "</a>"
	end if
	response.write "</th>"
	sqlvisual="select seqno,typename from visual_type where visible order by orders"
	rsvisual.open sqlvisual,conn,1,1
	TypeCount=0
	set rs=server.createobject("adodb.recordset")
	do while not rsvisual.eof
		sqlvisual="select top 1 seqno from visual_subtype where typeseq="&rsvisual("seqno")&" and visible order by orders"
		rs.open sqlvisual,conn,1,1
		if not rs.eof then
			response.write "<th valign=middle id=tabletitlelink>"
			if CurType=rsvisual("seqno") then
				response.write "<font color="&forum_body(8)&">"
				TypeName=rsvisual("TypeName")
			else
				minsubtype=rs(0)
				response.write "<a href=?shopid="&minsubtype&"0"&rsvisual("seqno")&"&sex="&sex&"&page=1>"
			end if
			response.write rsvisual("TypeName")
			if CurType=rsvisual("seqno") then
				response.write "</font>"
			else
				response.write "</a>"
			end if
			response.write "</th>"
			TypeCount=TypeCount+1
		end if
		rs.close
		rsvisual.movenext
	loop
	set rs=nothing
	rsvisual.close
	response.write "<th valign=middle id=tabletitlelink>"
	if CurType=96 then
		response.write "<font color="&forum_body(8)&">"
		TypeName="�����г�"
	else
		response.write "<a href=?shopid=196&sex="&sex&"&page=1>"
	end if
	response.write "�����г�"
	if CurType=96 then
		response.write "</font>"
	else
		response.write "</a>"
	end if
	response.write "</th>"
	response.write "<th valign=middle id=tabletitlelink>"
	if CurType=97 then
		response.write "<font color="&forum_body(8)&">"
		TypeName="�����"
	else
		response.write "<a href=?shopid=197>"
	end if
	response.write "�����"
	if CurType=97 then
		response.write "</font>"
	else
		response.write "</a>"
	end if
	response.write "</th>"
	response.write "</tr>"
	if CurType=95 then response.write "<form method=post action=?shopid=195&search_str="&request("search_str")&"&selgender="&request("selgender")&"&selprice_l="&request("selprice_l")&"&selprice_h="&request("selprice_h")&"&selperiod_l="&request("selperiod_l")&"&selperiod_h="&request("selperiod_h")&"&selflag="&request("selflag")&">"
	response.write "<tr height=25>"
	response.write "<td align=center valign=middle colspan="&(TypeCount+4)&" class=tablebody1>"
	if CurType=0 then
		if CurSubType=1 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵ��³�"
		else
			response.write "��<a href=?shopid=100&page=1>"
		end if
		response.write "�ҵ��³�"
		if CurSubType=1 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=2 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵĻ�ױ��"
		else
			response.write "��<a href=?shopid=200&page=1>"
		end if
		response.write "�ҵĻ�ױ��"
		if CurSubType=2 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=3 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵ����κ�"
		else
			response.write "��<a href=?shopid=300&page=1>"
		end if
		response.write "�ҵ����κ�"
		if CurSubType=3 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=4 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵ�������"
		else
			response.write "��<a href=?shopid=400&page=1>"
		end if
		response.write "�ҵ�������"
		if CurSubType=4 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=5 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵĻ���"
		else
			response.write "��<a href=?shopid=500&page=1>"
		end if
		response.write "�ҵĻ���"
		if CurSubType=5 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=6 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵ�����¨"
		else
			response.write "��<a href=?shopid=600&page=1>"
		end if
		response.write "�ҵ�����¨"
		if CurSubType=6 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
	elseif CurType=99 then
		if CurSubType=1 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�����ϼ�"
		else
			response.write "��<a href=?shopid=199&sex="&sex&"&page=1>"
		end if
		response.write "�����ϼ�"
		if CurSubType=1 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=2 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="������Ʒ"
		else
			response.write "��<a href=?shopid=299&sex="&sex&"&page=1>"
		end if
		response.write "������Ʒ"
		if CurSubType=2 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=3 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="VIPר��"
		else
			response.write "��<a href=?shopid=399&sex="&sex&"&page=1>"
		end if
		response.write "VIPר��"
		if CurSubType=3 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=4 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ؼ���Ʒ"
		else
			response.write "��<a href=?shopid=499&sex="&sex&"&page=1>"
		end if
		response.write "�ؼ���Ʒ"
		if CurSubType=4 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
	elseif CurType=96 then
		response.write "&nbsp;&nbsp;"
		if CurSubType=1 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="��װ��"
		else
			response.write "��<a href=?shopid=196&sex="&sex&"&page=1>"
		end if
		response.write "��װ��"
		if CurSubType=1 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=2 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="��ױ��"
		else
			response.write "��<a href=?shopid=296&sex="&sex&"&page=1>"
		end if
		response.write "��ױ��"
		if CurSubType=2 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=3 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="������"
		else
			response.write "��<a href=?shopid=396&sex="&sex&"&page=1>"
		end if
		response.write "������"
		if CurSubType=3 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=4 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="������"
		else
			response.write "��<a href=?shopid=496&sex="&sex&"&page=1>"
		end if
		response.write "������"
		if CurSubType=4 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
	elseif CurType=97 then
		response.write "&nbsp;&nbsp;"
		if CurSubType=1 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="����������"
		else
			response.write "��<a href=?shopid=197>"
		end if
		response.write "����������"
		if CurSubType=1 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=2 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="����������"
		else
			response.write "��<a href=?shopid=297>"
		end if
		response.write "����������"
		if CurSubType=2 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=3 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�յ�������"
		else
			response.write "��<a href=?shopid=397>"
		end if
		response.write "�յ�������"
		if CurSubType=3 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=4 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="ȷ�ϵ�����"
		else
			response.write "��<a href=?shopid=497>"
		end if
		response.write "ȷ�ϵ�����"
		if CurSubType=4 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=5 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�ҵ������ಾ"
		else
			response.write "��<a href=?shopid=597>"
		end if
		response.write "�ҵ������ಾ"
		if CurSubType=5 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=6 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="��ҵ������ಾ"
		else
			response.write "��<a href=?shopid=697>"
		end if
		response.write "��ҵ������ಾ"
		if CurSubType=6 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
	elseif CurType=98 then
		response.write "&nbsp;&nbsp;"
		if CurSubType=1 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="����"
		else
			response.write "��<a href=?shopid=198&sex="&sex&"&page=1>"
		end if
		response.write "����"
		if CurSubType=1 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=2 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="������Ʒ"
		else
			response.write "��<a href=?shopid=298&sex="&sex&"&page=1>"
		end if
		response.write "������Ʒ"
		if CurSubType=2 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=3 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�����̵�"
		else
			response.write "��<a href=?shopid=398&sex="&sex&"&page=1>"
		end if
		response.write "�����̵�"
		if CurSubType=3 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=4 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�����г�"
		else
			response.write "��<a href=?shopid=498&sex="&sex&"&page=1>"
		end if
		response.write "�����г�"
		if CurSubType=4 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
		response.write "&nbsp;&nbsp;"
		if CurSubType=5 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�����"
		else
			response.write "��<a href=?shopid=598&sex="&sex&"&page=1>"
		end if
		response.write "�����"
		if CurSubType=5 then
			response.write "</font>��"
		else
			response.write "</a>��"
		end if
response.write "&nbsp;&nbsp;"
if CurSubType=6 then
			response.write "��<font color="&forum_body(8)&">"
			SubTypeName="�������"
		else
			response.write "<font color=#ff00ff>��<a href=z_Visual_photo2.asp>"
		end if
		response.write "<font color=#ff00ff>�������</font>"
		if CurSubType=6 then
			response.write "</font>��"
		else
			response.write "</a>��</font>"
		end if
	elseif CurType=95 then
		response.write "�������"
		response.write "<input type=radio name=sortobj value=""id"""
		if isnull(request("sortobj")) or request("sortobj")="" then
			response.write " checked"
		else
			if not request("sortobj")<>"id" or (request("sortobj")<>"dateandtime" and request("sortobj")<>"price1" and request("sortobj")<>"price2" and request("sortobj")<>"period" and request("sortobj")<>"flag") then response.write " checked"
		end if
		response.write ">��Ʒ����&nbsp;"
		response.write "<input type=radio name=sortobj value=""dateandtime"""
		if not request("sortobj")<>"dateandtime" then response.write " checked"
		response.write ">����ʱ��&nbsp;"
		response.write "<input type=radio name=sortobj value=""price1"""
		if not request("sortobj")<>"price1" then response.write " checked"
		response.write ">���&nbsp;"
		response.write "<input type=radio name=sortobj value=""price2"""
		if not request("sortobj")<>"price2" then response.write " checked"
		response.write ">VIP��&nbsp;"
		response.write "<input type=radio name=sortobj value=""period"""
		if not request("sortobj")<>"period" then response.write " checked"
		response.write ">��Ч��&nbsp;"
		response.write "<input type=radio name=sortobj value=""flag"""
		if not request("sortobj")<>"flag" then response.write " checked"
		response.write ">�޹���&nbsp;"
		response.write "����ʽ��"
		response.write "<input type=radio name=sorttype value=""asc"""
		if not request("sorttype")<>"asc" then response.write " checked"
		response.write ">����&nbsp;"
		response.write "<input type=radio name=sorttype value=""desc"""
		if isnull(request("sorttype")) or request("sorttype")="" then
			response.write " checked"
		else
			if not request("sorttype")<>"desc" or request("sorttype")<>"asc" then response.write " checked"
		end if
		response.write ">����&nbsp;"
		response.write "<input type=""submit"" name=""submit"" value=""����"">"
	else
		sqlvisual="select seqno,subtypename from visual_subtype where typeseq="&CurType&" and visible order by orders"
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
			if CurSubType=rsvisual("seqno") then
				response.write "��<font color="&forum_body(8)&">"
				SubTypeName=rsvisual("subtypename")
			else
				response.write "��<a href=?shopid="&rsvisual("seqno")&right("00"&CurType,2)&"&sex="&sex&"&page=1>"
			end if
			response.write rsvisual("subtypename")
			if CurSubType=rsvisual("seqno") then
				response.write "</font>��"
			else
				response.write "</a>��"
			end if
			rsvisual.movenext
			if not rsvisual.eof then response.write "&nbsp;&nbsp;"
		loop
		rsvisual.close
	end if
	response.write "</td>"
	response.write "</tr>"
	if CurType=95 then response.write "</form>"
	response.write "</table>"
	response.write "<br>"
	if CurType<>0 and CurType<>96 and CurType<>97 and CurType<>98 then
		response.write "<table cellpadding=2 cellspacing=1 align=center class=tableborder1>"
		response.write "<tr height=25>"
		response.write "<th valign=middle id=tabletitlelink>����������Ʒ</th>"
		response.write "</tr>"
		response.write "<form method=post action=?shopid=195>"
		response.write "<tr height=30>"
		response.write "<td align=center valign=middle width=100% class=tablebody1>"
		response.write "��Ʒ���ƣ�<input type=""text"" name=""search_str"" size=""10"" maxlength=""20"">&nbsp;"
		response.write "<select name=""selgender""><option value=0 selected>�Ա���</option><option value=1>��</option><option value=2>Ů</option><option value=3>��Ů����</option></select>&nbsp;"
		response.write "<input type=""text"" name=""selPrice_L"" size=""5"" maxlength=""5"">�ܼ۸��<input type=""text"" name=""selPrice_H"" size=""5"" maxlength=""5"">&nbsp;"
		response.write "<input type=""text"" name=""selPeriod_L"" size=""5"" maxlength=""5"">����Ч�ڡ�<input type=""text"" name=""selPeriod_H"" size=""5"" maxlength=""5"">&nbsp;"
		response.write "<select name=""selflag""><option value=0 selected>�޹��߲���</option><option value=1>���л�Ա</option><option value=2>��������</option><option value=3>��������</option><option value=4>����Ա</option></select>&nbsp;"
		response.write "<input type=""submit"" name=""submit"" value=""��ʼ����"">"
		response.write "</td>"
		response.write "</tr>"
		response.write "</form>"
		response.write "</table>"
		response.write "<br>"
	end if
	response.write "<a name=visualtop>"
	response.write "<table cellpadding=2 cellspacing=1 align=center class=tableborder1>"
	response.write "<tr height=25>"
	if CurType=96 then
		response.write "<th valign=middle id=tabletitlelink colspan=2><marquee scrolldelay=100 scrollamount=2>�����̵�ר�ž�Ӫ�ɻ����۸���ˣ����Ƿ�����������Ʒ<font color="&forum_body(8)&">�����Դ�</font>����������ʱС��Ӵ! </marquee></th>"
	elseif CurType=97 then
		response.write "<th valign=middle id=tabletitlelink colspan=2><marquee scrolldelay=100 scrollamount=2>����������࣬����ı����ճ�ʹ�õ�����ֻҪװ������Ҫ����Ʒ�����ñ��漴�ɺ�Ӱ��д��! </marquee></th>"
	elseif CurType=98 then
		response.write "<th valign=middle id=tabletitlelink colspan=2>"&forum_info(0)&" -- �������ϵͳ</th>"
	else
		response.write "<th valign=middle id=tabletitlelink colspan=2><marquee scrolldelay=100 scrollamount=2>���ͼƬ�����Դ������£��Դ���ǵá�<font color="&forum_body(8)&">����</font>��Ӵ! </marquee></th>"
	end if
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td align=center valign=middle class=tablebody2 width=160>"
	response.write "<table border=0 cellspacing=1 cellpadding=0 width=140 bgcolor="&forum_body(27)&">"
	response.write "<tr>"
	response.write "<script language=JavaScript>"
	response.write "var cookieSrc=document.cookie;"
	response.write "</script>"
	response.write "<td colspan=2 class=tablebody1 width=140 height=226 id=visualshow ImagePath="""&PicPath&""" enableLayers="""&myLayerStr&""" usergender="""&v_mysex&""" baseName=""images/img_visual/blank.gif"" "
	for i=0 to ubound(myvisualsplit)
		if not isnull(myvisualsplit(i)) and trim(myvisualsplit(i))<>"" then
			response.write "layer"&(i+1)&"="&myvisualsplit(i)&" "
		end if
	next
	response.write "style=behavior:url(inc/z_show.htc) cookies=cookieSrc setCookies=""document.cookie=this.cookies"" userid="&userid&" localpic="&LocalPic&">"
	response.write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
	if CurType<>97 then
		response.write "<tr>"
		response.write "<td class=TableBody2 height=27 align=center><a href=z_visual_save.asp>����</a>&nbsp;<a href=javascript:initshow()>���</a>&nbsp;<a href=javascript:deleteshow()>�ָ�</a></td>"
		response.write "</tr>"
	else
		response.write "<tr>"
		response.write "<td class=TableBody2 height=27 align=center><font color=gray>����</font>&nbsp;<a href=javascript:initshow()>���</a>&nbsp;<a href=javascript:deleteshow()>�ָ�</a></td>"
		response.write "</tr>"
	end if
	response.write "<tr>"
	response.write "<td class=TableBody2 height=27 align=center><a href=?shopid=197>��Ӱ</a>&nbsp;<a href=z_visual_photo.asp?action=photome onclick=""if(confirm('����д��ķ�����"
	if isnull(v_myvip) or v_myvip<>1 then
		response.write PhotoPrice1
	else
		response.write PhotoPrice2
	end if
	response.write "Ԫ�����������ֽ�"&mymoney&"Ԫ���Ƿ�ȷ������?')){return true;} return false;"">д��</a>&nbsp;<font color=gray title=�¸��汾�ṩ�˹���>����</font></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table name=takeonrecord align=center width=140 cellpadding=0 cellspacing=0>"
	response.write "<tr>"
	response.write "<td width=70 align=center><img style=cursor:hand height=70 name=""IMG"" src=""images/img_visual/blank.gif"" width=70 border=0 itemgender="&v_mysex&"></td>"
	response.write "<td width=70 align=center><img style=cursor:hand height=70 name=""IMG"" src=""images/img_visual/blank.gif"" width=70 border=0 itemgender="&v_mysex&"></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table width=140 border=0 cellspacing=0 cellpadding=0>"
	response.write "<tr height=25>"
	response.write "<td width=80 class=tablebody2>&nbsp;&nbsp;[�Դ���¼]</td>"
	response.write "<td name=btnMoveLeft width=30 valign=middle align=center class=tablebody2><img src=""images/img_visual/left.gif"" style=""CURSOR: hand; display:none""></td>"
	response.write "<td name=btnMoveRight width=30 valign=middle align=center class=tablebody2><img src=""images/img_visual/right.gif"" style=""CURSOR: hand; display:none""></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "</td>"
	response.write "</tr>"
	response.write "<script language=javascript>"
	response.write "function deleteshow(){"
	response.write "document.cookie = ""myshow_"&userid&"="" + escape("""") + ""; path="&cookiepath&""";"
	response.write "location.reload();"
	response.write "}"
	response.write "function initshow(){"
 if v_mysex=1 then
  response.write "document.cookie = ""myshow_"&userid&"="" + escape('||||||14|13|12||11||10|9||||8|||||||') + ""; path="&cookiepath&""";"
 else
  response.write "document.cookie = ""myshow_"&userid&"="" + escape('||||||7|6|5||4||3|2||||1|||||||') + ""; path="&cookiepath&""";"
 end if
 response.write "location.reload();"
	response.write "}"
	response.write "</script>"
	response.write "<tr height=35 valign=middle>"
	response.write "<td colspan=2 class=tablebody2 align=center><input type=button style=""cursor:hand;"" onclick=""JavaScript:openScript('z_visual_cartbag.asp',500,440)"" value=�鿴�����></td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td width=50% align=right class=tablebody1>�Ա�&nbsp;</td>"
	response.write "<td align=left class=tablebody1>&nbsp;"
	if v_mysex=0 then
		response.write "Ů"
	else
		response.write "��"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td width=50% align=right class=tablebody1>���ԣ�&nbsp;</td>"
	response.write "<td align=left class=tablebody1>&nbsp;"
	if v_myvip<>1 then
		response.write "��ͨ��Ա"
	else
		response.write "<font color="&forum_body(8)&"><b>VIP ��Ա</b></font>"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td width=50% align=right class=tablebody1>�ȼ���&nbsp;</td>"
	response.write "<td align=left class=tablebody1>&nbsp;"
	if master then
		response.write "����Ա"
	elseif superboardmaster then
		response.write "��������"
	elseif boardmaster then
		response.write "����"
	else
		response.write "��ͨ��Ա"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td width=50% align=right class=tablebody1>�ֽ�&nbsp;</td>"
	response.write "<td align=left class=tablebody1>&nbsp;"&mymoney&"</td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td colspan=2 class=tablebody1 align=center><font style=""cursor:hand"" onclick=""JavaScript:openScript('z_visual_record.asp?menu=buy',500,400)"" color=blue>�鿴�����¼</font></td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td colspan=2 class=tablebody1 align=center><font style=""cursor:hand"" onclick=""JavaScript:openScript('z_visual_record.asp?menu=send',500,400)"" color=blue>�鿴���ͼ�¼</font></td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td colspan=2 class=tablebody1 align=center><font style=""cursor:hand"" onclick=""JavaScript:openScript('z_visual_record.asp?menu=recv',500,400)"" color=blue>�鿴������¼</font></td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td colspan=2 class=tablebody1 align=center><font style=""cursor:hand"" onclick=""JavaScript:openScript('z_visual_record.asp?menu=sale',500,400)"" color=blue>�鿴���ۼ�¼</font></td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td colspan=2 class=tablebody1 align=center><a href="&scriptname&"><font color=blue>[ʹ�ð���]</font></a></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "</td>"
	response.write "<td align=center valign=top class=tablebody1>"
	if CurType<>98 and CurType<>97 then
		if CurType=0 then
			if shopid=100 then
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=1 and my.username='"&membername&"' order by my.id desc"
			elseif shopid=200 then
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=2 and my.username='"&membername&"' order by my.id desc"
			elseif shopid=300 then
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=3 and my.username='"&membername&"' order by my.id desc"
			elseif shopid=400 then
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=4 and my.username='"&membername&"' order by my.id desc"
			elseif shopid=500 then
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0,saleprice from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=5 and my.username='"&membername&"' order by my.id desc"
			else
				sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.type=0 and my.username='"&membername&"' order by my.id desc"
			end if
		elseif CurType=99 then
			if shopid=199 then
				if sex="1" or sex="0" then
					sqlvisual="select top 100 * from visual_items where (sex="&sex&" or sex=2) and flag>0 and type>0 order by DateAndTime desc"
				else
					sqlvisual="select top 100 * from visual_items where flag>0 and type>0 order by DateAndTime desc"
				end if
			elseif shopid=299 then
				if sex="1" or sex="0" then
					sqlvisual="select top 50 * from visual_items where (sex="&sex&" or sex=2) and flag>0 and type>0 order by totalsales desc, id desc"
				else
					sqlvisual="select top 50 * from visual_items where flag>0 and type>0 order by totalsales desc, id desc"
				end if
			elseif shopid=399 then
				if sex="1" or sex="0" then
					sqlvisual="select * from visual_items where (sex="&sex&" or sex=2) and flag>0 and type>0 and vip order by id desc"
				else
					sqlvisual="select * from visual_items where flag>0 and vip and type>0 order by id desc"
				end if
			else
				if sex="1" or sex="0" then
					sqlvisual="select * from visual_items where (sex="&sex&" or sex=2) and flag>0 and type>0 and price1<="&PriceMin&" order by id desc"
				else
					sqlvisual="select * from visual_items where flag>0 and price1<="&PriceMin&" and type>0 order by id desc"
				end if
			end if
		elseif CurType=96 then
			if shopid=196 then
				if sex="1" or sex="0" then
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where (v.sex="&sex&" or v.sex=2) and m.type=5 and v.type mod 100=1 order by m.id desc"
				else
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where m.type=5 and v.type mod 100=1 order by m.id desc"
				end if
			elseif shopid=296 then
				if sex="1" or sex="0" then
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where (v.sex="&sex&" or v.sex=2) and m.type=5 and v.type mod 100=2 order by m.id desc"
				else
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where m.type=5 and v.type mod 100=2 order by m.id desc"
				end if
			elseif shopid=396 then
				if sex="1" or sex="0" then
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where (v.sex="&sex&" or v.sex=2) and m.type=5 and v.type mod 100=3 order by m.id desc"
				else
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where m.type=5 and v.type mod 100=3 order by m.id desc"
				end if
			else
				if sex="1" or sex="0" then
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where (v.sex="&sex&" or v.sex=2) and m.type=5 and v.type mod 100<>1 and v.type mod 100<>2 and v.type mod 100<>3 and v.type mod 100<>4 order by m.id desc"
				else
					sqlvisual="select m.ItemID as ID,v.Name,v.sex,v.content,v.price1,v.price2,v.period as period0,m.username,m.period,m.saleprice,m.fixdate,m.ID as ID1 from visual_myitems m inner join visual_items v on m.itemid=v.id where m.type=5 and v.type mod 100<>1 and v.type mod 100<>2 and v.type mod 100<>3 order by m.id desc"
				end if
			end if
		elseif CurType=95 then
			sqlvisual="select * from visual_items where flag>0 and type>0"
			if not (isnull(request("search_str")) or request("search_str")="") then
				sqlvisual=sqlvisual&" and name like '%"&checkstr(trim(request("search_str")))&"%'"
				searchmsg="����[<font color="&forum_body(8)&">"&checkstr(trim(request("search_str")))&"</font>]"
			else
				searchmsg="����[]"
			end if
			searchmsg=searchmsg&"���Ա�"
			if not (isnull(request("selgender")) or request("selgender")="") then
				select case request("selgender")
				case "1"
					sqlvisual=sqlvisual&" and sex=1"
					searchmsg=searchmsg&"<font color="&forum_body(8)&">��</font>"
				case "2"
					sqlvisual=sqlvisual&" and sex=0"
					searchmsg=searchmsg&"<font color="&forum_body(8)&">Ů</font>"
				case "3"
					sqlvisual=sqlvisual&" and sex=2"
					searchmsg=searchmsg&"<font color="&forum_body(8)&">��Ů����</font>"
				case else
					searchmsg=searchmsg&"<font color="&forum_body(8)&">����</font>"
				end select
			else
				searchmsg=searchmsg&"<font color="&forum_body(8)&">����</font>"
			end if
			dim hasFlag
			hasFlag=false
			searchmsg=searchmsg&"���۸�"
			if not (isnull(request("selprice_l")) or request("selprice_l")="") then
				if isInteger(request("selprice_l")) then
					if v_myVip<>1 then
						sqlvisual=sqlvisual&" and price1>="&request("selprice_l")
					else
						sqlvisual=sqlvisual&" and price2>="&request("selprice_l")
					end if
					searchmsg=searchmsg&"<font color="&forum_body(8)&">"&request("selprice_l")&"Ԫ����</font>"
					hasFlag=true
				end if
			end if
			if not (isnull(request("selprice_h")) or request("selprice_h")="") then
				if isInteger(request("selprice_h")) then
					if v_myVip<>1 then
						sqlvisual=sqlvisual&" and price1<="&request("selprice_h")
					else
						sqlvisual=sqlvisual&" and price2<="&request("selprice_h")
					end if
					searchmsg=searchmsg&"<font color="&forum_body(8)&">"&request("selprice_h")&"Ԫ����</font>"
					hasFlag=true
				end if
			end if
			if not hasFlag then searchmsg=searchmsg&"<font color="&forum_body(8)&">����</font>"
			hasFlag=false
			searchmsg=searchmsg&"����Ч��"
			if not (isnull(request("selperiod_l")) or request("selperiod_l")="") then
				if isInteger(request("selperiod_l")) then
					sqlvisual=sqlvisual&" and period>="&request("selperiod_l")
					searchmsg=searchmsg&"<font color="&forum_body(8)&">"&request("selperiod_l")&"������</font>"
					hasFlag=true
				end if
			end if
			if not (isnull(request("selperiod_h")) or request("selperiod_h")="") then
				if isInteger(request("selperiod_h")) then
					sqlvisual=sqlvisual&" and period<="&request("selperiod_h")
					searchmsg=searchmsg&"<font color="&forum_body(8)&">"&request("selperiod_h")&"������</font>"
					hasFlag=true
				end if
			end if
			if not hasFlag then searchmsg=searchmsg&"<font color="&forum_body(8)&">����</font>"
			searchmsg=searchmsg&"��"
			if not (isnull(request("selflag")) or request("selflag")="") then
				select case request("selflag")
				case "1"
					sqlvisual=sqlvisual&" and flag>=1"
					searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">���л�Ա</font>"
				case "2"
					sqlvisual=sqlvisual&" and flag>=2"
					searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">��������</font>"
				case "3"
					sqlvisual=sqlvisual&" and flag>=3"
					searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">��������</font>"
				case "4"
					sqlvisual=sqlvisual&" and flag>=4"
					searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">����Ա</font>"
				case else
					searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">����</font>"
				end select
			else
				searchmsg=searchmsg&"�޹���<font color="&forum_body(8)&">����</font>"
			end if
			sqlvisual=sqlvisual&" order by "
			if isnull(request("sortobj")) or request("sortobj")="" then
				sqlvisual=sqlvisual&"id"
			else
				if request("sortobj")<>"dateandtime" and request("sortobj")<>"price1" and request("sortobj")<>"price2" and request("sortobj")<>"period" and request("sortobj")<>"flag" then
					sqlvisual=sqlvisual&"id"
				else
					sqlvisual=sqlvisual&request("sortobj")
				end if
			end if
			if isnull(request("sorttype")) or request("sorttype")="" then
				sqlvisual=sqlvisual&" desc"
			else
				if request("sorttype")<>"asc" then
					sqlvisual=sqlvisual&" desc"
				else
					sqlvisual=sqlvisual&" asc"
				end if
			end if
			searchmsg=searchmsg&"��"
		elseif CurType<>98 and CurType<>97  then
			if sex="1" or sex="0" then
				sqlvisual="select * from visual_items where type="&shopid&" and (sex="&sex&" or sex=2) and flag>0 order by id desc"
			else
				sqlvisual="select * from visual_items where type="&shopid&" and flag>0 order by id desc"
			end if
		end if
		rsvisual.open sqlvisual,conn,1,1
		totalrec=rsvisual.recordcount
		if totalrec mod 10=0 then
			pagecount=totalrec \ 10
		else
			pagecount=totalrec \ 10+1
		end if
		curcount=0
		response.write "<table id=visualshowGoods border=0 cellspacing=1 cellpadding=0 bgcolor="&Forum_Body(27)&">"
		response.write "<tr height=25 valign=middle>"
		response.write "<th valign=middle>"
		if CurType=95 then
			response.write searchmsg&"�������<font color="&forum_body(8)&">"&totalrec&"</font>��"
		else
			response.write TypeName&"&nbsp;-&nbsp;"&SubTypeName
			if CurType<>0 then
				response.write "&nbsp;&nbsp;(&nbsp;<a href=?shopid="&shopid&"&sex=1>��</a>&nbsp;<a href=?shopid="&shopid&"&sex=0>Ů</a>&nbsp;<a href=?shopid="&shopid&">ȫ��</a>&nbsp;)"
			end if
		end if
		response.write "</th>"
		response.write "</tr>"
		if not rsvisual.eof then
			if curpage > pagecount then curpage = pagecount
			if curpage < 1 then curpage=1
			Rsvisual.PageSize=10
			Rsvisual.AbsolutePage=CurPage
			do while not rsvisual.eof and curcount<10
				isHas=not (conn.execute("select top 1 itemid from visual_myitems where itemid="&rsvisual("id")&" and username='"&membername&"'").eof)
				curcount=curcount+1
				if curcount mod 2=1 then
					response.write "<tr valign=middle>"
					response.write "<td class=tablebody1 align=center>"
					response.write "<table width=560 border=0 cellspacing=0 cellpadding=0>"
					response.write "<tr>"
				else
					response.write "<td width=14>&nbsp;</td>"
				end if
				response.write "<td width=282>"
				response.write "<table width=282 border=0 cellspacing=2 cellpadding=0>"
				response.write "<tr valign=middle>"
				response.write "<td rowspan=5 class=tablebody2 width=84 height=84>"
      	if localpic=1 then
      		response.write "<img src="""&PicPath&"0/"&right("0000"&rsvisual("id"),4)&".GIF"" "
      	else
      		response.write "<img src="""&PicPath&right("00000000"&rsvisual("id"),8)&"/00/00/"" "
      	end if
    		if shopid=600 then
    			response.write "alt=""������Ʒ����ʹ��"" width=""84"" height=""84"" border=""0"" style=""CURSOR:help"">"
        elseif CurType=96 then
          response.write "alt=""������Ʒ�����Դ�"" width=""84"" height=""84"" border=""0"" style=""CURSOR:help"">"
        elseif CurType<>0 and isHas then
          response.write "alt=""������Ʒ�̵겻���Դ�"" width=""84"" height=""84"" border=""0"" style=""CURSOR:help"">"
        else
        	response.write "alt=""����������Դ�"" width=""84"" height=""84"" border=""0"" itemgender="""&rsvisual("sex")&""" itemno="""&rsvisual("id")&""" layerno="""&rsvisual("content")&""" style=""CURSOR:hand"" onclick=""visualshow.clickDress()"">"
        end if
				response.write "</td>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
    		response.write rsvisual("name")
      	response.write "("
      	if rsvisual("sex")=1 then
      		response.write "��"
      	elseif rsvisual("sex")=0 then
      		response.write "Ů"
      	else
      		response.write "����"
        end if
        response.write ")"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
      	if CurType=0 then
      		if shopid=500 then
        		response.write "���ۼۣ�"
      			response.write rsvisual("saleprice")&" Ԫ"
        	else
        		response.write "����ۣ�"
        		if rsvisual("price")<=0 then
        			response.write "<font color="&forum_body(8)&">���</font>"
        		else
        			response.write rsvisual("price")&" Ԫ"
        		end if
        	end if
      	elseif CurType=96 then
      		response.write "�̵�ۣ�"
      		if v_myvip<>1 then
        		if rsvisual("price1")<=0 then
        			response.write "<font color="&forum_body(8)&">���</font>"
        		else
        			response.write rsvisual("price1")&" Ԫ"
        		end if
        	else
        		if rsvisual("price2")<=0 then
        			response.write "<font color="&forum_body(8)&">���</font>"
        		else
        			response.write rsvisual("price2")&" Ԫ"
        		end if
        	end if
    		else
      		response.write "�ꡡ�ۣ�"
      		if rsvisual("price1")<=0 then
      			response.write "<font color="&forum_body(8)&">���</font>"
      		else
      			response.write rsvisual("price1")&" Ԫ"
      		end if
      	end if
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
      	if CurType=0 then
      		if shopid=500 then
      			response.write "�ɱ��ۣ�"
      		else
      			response.write "��ǰ�ۣ�"
      		end if
      		if v_myvip<>1 then
        		if rsvisual("price1")<=0 then
        			response.write "<font color="&forum_body(8)&">���</font>"
        		else
        			response.write rsvisual("price1")&" Ԫ"
        		end if
        	else
        		if rsvisual("price2")<=0 then
        			response.write "<font color="&forum_body(8)&">���</font>"
        		else
        			response.write rsvisual("price2")&" Ԫ"
        		end if
        	end if
      	elseif CurType=96 then
      		response.write "���ּۣ�"
    			response.write "<font color="&forum_body(8)&">"&rsvisual("saleprice")&"</font> Ԫ"
    		else
      		response.write "VIP �ۣ�"
      		if rsvisual("price2")<=0 then
      			response.write "<font color="&forum_body(8)&">���</font>"
      		else
      			response.write "<font color="&forum_body(8)&">"&rsvisual("price2")&"</font> Ԫ"
      		end if
      	end if
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
    		if CurType=0 then
      		response.write "��Ч�ڣ�"
      		if rsvisual("period")=0 then
      			response.write "<font color="&forum_body(8)&">��Զ��Ч</font>"
    			elseif datediff("d",rsvisual("fixdate"),now())>rsvisual("period") then
      			response.write "<font color="&forum_body(8)&">�Ѿ�����</font>"
    			else
    				response.write "��ʣ "&(rsvisual("period")-datediff("d",rsvisual("fixdate"),now()))&" ��"
    			end if
    		elseif CurType=96 then
      		response.write "��Ʒ�ڣ�"
      		if rsvisual("period0")<=0 then
      			response.write "<font color="&forum_body(8)&">����</font>"
      		else
      			response.write rsvisual("period0")&" ��"
      		end if
    		else
      		response.write "��Ч�ڣ�"
      		if rsvisual("period")<=0 then
      			response.write "<font color="&forum_body(8)&">����</font>"
      		else
      			response.write rsvisual("period")&" ��"
      		end if
      	end if
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
    		if CurType=0 then
        	response.write "ӵ���գ�"
        	response.write formatdatetime(rsvisual("adddate"),2)
        elseif CurType=96 then
      		response.write "�����ڣ�"
      		if rsvisual("period")<=0 then
      			response.write "<font color="&forum_body(8)&">��Զ��Ч</font>"
    			elseif datediff("d",rsvisual("fixdate"),now())>rsvisual("period") then
      			response.write "<font color="&forum_body(8)&">�Ѿ�����</font>"
      		else
      			response.write "��ʣ "&(rsvisual("period")-datediff("d",rsvisual("fixdate"),now()))&" ��"
      		end if
      	else
        	response.write "�⡡�棺"
      		if rsvisual("quantity")<=0 or isnull(rsvisual("quantity")) then
      			response.write "<font color="&forum_body(8)&">�ϻ�</font>"
      		else
      			response.write rsvisual("quantity")&" ��"
      		end if
      	end if
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20 align=center>"
    		if CurType=0 then
    			if rsvisual("fromuser")=rsvisual("username") then
    				response.write "<font color=blue alt=��Դ>�����̵�</font>"
    			else
    				response.write "<font color=blue alt=��Դ>"&rsvisual("fromuser")&"</font>"
    			end if
    		elseif CurType=96 then
  				response.write "<font color=blue>������Ʒ</font>"
  			else
      		if rsvisual("vip") then
      			response.write "<font color="&forum_body(8)&">VIP��Ʒ</font>"
      		else
      			response.write "��ͨ��Ʒ"
      		end if
      	end if
				response.write "</td>"
				response.write "<td class=tablebody2 height=20>&nbsp;&nbsp;"
    		if CurType=0 then
        	response.write "�����գ�"
        	if isnull(rsvisual("fixdate")) or (rsvisual("fixdate")=rsvisual("adddate") and rsvisual("period")=rsvisual("period0")) then
        		response.write "��δ����"
        	else
        		response.write "<font color="&forum_body(8)&">"&formatdatetime(rsvisual("fixdate"),2)&"</font>"
        	end if
        elseif CurType=96 then
        	response.write "�����ߣ�"
      		response.write "<a href=dispuser.asp?name="&rsvisual("username")&" target=_blank><font color=blue>"&rsvisual("username")&"</font></a>"
      	else
        	response.write "�޹��ߣ�"
        	select case rsvisual("flag")
        	case 1
        		response.write "���л�Ա"
        	case 2
        		response.write "<font color="&forum_body(8)&">��������</font>"
        	case 3
        		response.write "<font color="&forum_body(8)&">��������</font>"
        	case 4
        		response.write "<font color="&forum_body(8)&">����Ա</font>"
        	end select
        end if
				response.write "</td>"
				response.write "</tr>"
			  if CurType=96 then
				  response.write "<form name=form1_"&rsvisual("id1")&" action=z_visual_buy.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"><input type=hidden name=fromuser value=""""></form>"
				  response.write "<form name=form2_"&rsvisual("id1")&" action=z_visual_send.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"><input type=hidden name=myself value=0><input type=hidden name=fromuser value=""""></form>"
				else
				  response.write "<form name=form1_"&rsvisual("id")&" action=z_visual_buy.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"><input type=hidden name=fromuser value=""""></form>"
				  response.write "<form name=form2_"&rsvisual("id")&" action=z_visual_send.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"><input type=hidden name=myself value=0><input type=hidden name=fromuser value=""""></form>"
				end if
			  response.write "<form name=form3_"&rsvisual("id")&" action=z_visual_fix.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"></form>"
			  response.write "<form name=form4_"&rsvisual("id")&" action=z_visual_sale.asp method=get target=openScript><input type=hidden name=itemid value="&rsvisual("id")&"><input type=hidden name=stopsale value=0></form>"
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody2 height=20 align=center>��"
    		if CurType=0 then
    			if rsvisual("period")>0 then
    				response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:openScript('about:blank',500,400);form3_"&rsvisual("id")&".submit()"">����</font></b>"
    			else
      			response.write "<b><font color=gray style=""cursor:hand"" alt='��Զ��Ч����Ʒ���뷭�£�'>����</font></b>"
      		end if
    		elseif CurType=96 then
  				response.write "<b><font color=gray style=""cursor:hand"" alt='������Ʒ�����Դ���'>�Դ�</font></b>"
  			else
    			if not isHas then
    				response.write "<b><font color=blue style=""cursor:hand"" alt=""�Դ�"" itemgender="""&rsvisual("sex")&""" itemno="""&rsvisual("id")&""" layerno="""&rsvisual("content")&""" onclick=""visualshow.clickDress()"">�Դ�</font></b>"
    			else
    				response.write "<b><font color=gray style=""cursor:hand"" alt='�����Ʒ���Ѿ����ˣ��뵽���Ĵ�������Դ���'>�Դ�</font></b>"
    			end if
    		end if
      	response.write "��"
				response.write "</td>"
				response.write "<td class=tablebody2 height=20 align=center>��"
    		if CurType=0 then
    			if instr(1,"|"&v_myvisual,"|"&rsvisual("id")&".")<=0 then
      			response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:form2_"&rsvisual("id")&".myself.value=1;openScript('about:blank',500,400);form2_"&rsvisual("id")&".submit()"">����</font></b>"
      		else
      			response.write "<b><font color=gray style=""cursor:hand"" alt='����ʹ�õ���Ʒ�������ͣ�'>����</font></b>"
      		end if
        	response.write "��&nbsp;&nbsp;��"
    			if shopid=500 then
      			response.write "<b><font color=blue style=""cursor:hand"" alt=""ͣ��"" onclick=""javascript:form4_"&rsvisual("id")&".stopsale.value=1;openScript('about:blank',500,400);form4_"&rsvisual("id")&".submit()"">ͣ��</font></b>"
      		else
      			if instr(1,"|"&v_myvisual,"|"&rsvisual("id")&".")<=0 then
        			response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:form4_"&rsvisual("id")&".stopsale.value=0;openScript('about:blank',500,400);form4_"&rsvisual("id")&".submit()"">����</font></b>"
        		else
        			response.write "<b><font color=gray style=""cursor:hand"" alt='����ʹ�õ���Ʒ���ܳ��ۣ�'>����</font></b>"
        		end if
      		end if
      	else
      		if CurType=96 then
      			if lcase(trim(rsvisual("username")))=lcase(trim(membername)) then
      				response.write "<b><font color=gray style=""cursor:hand"" alt='���������۵���Ʒ�����Լ����ܹ���'>����</font></b>"
      			elseif not isHas then
        			response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:if(confirm('��ȷ�������������������Ʒ��?')){form1_"&rsvisual("id1")&".fromuser.value='"&rsvisual("username")&"';openScript('about:blank',500,400);form1_"&rsvisual("id1")&".submit();return true;} return false;"">����</font></b>"
        		else
      				response.write "<b><font color=gray style=""cursor:hand"" alt='�����Ʒ���Ѿ����ˣ������ٴι���'>����</font></b>"
      			end if
    			else
      			if not isHas then
        			response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:form1_"&rsvisual("id")&".fromuser.value='';openScript('about:blank',500,400);form1_"&rsvisual("id")&".submit()"">����</font></b>"
        		else
      				response.write "<b><font color=gray style=""cursor:hand"" alt='�����Ʒ���Ѿ����ˣ������ٴι���'>����</font></b>"
      			end if
      		end if
        	response.write "��&nbsp;&nbsp;��"
    			if CurType=96 then
      			if lcase(trim(rsvisual("username")))=lcase(trim(membername)) then
      				response.write "<b><font color=gray style=""cursor:hand"" alt='���������۵���Ʒ��������������ȡ�����ۣ�'>����</font></b>"
      			else
      				response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:form2_"&rsvisual("id1")&".myself.value=0;form2_"&rsvisual("id1")&".fromuser.value='"&rsvisual("username")&"';openScript('about:blank',500,400);form2_"&rsvisual("id1")&".submit()"">����</font></b>"
      			end if
      		else
    				response.write "<b><font color=blue style=""cursor:hand"" alt=""����"" onclick=""javascript:form2_"&rsvisual("id")&".myself.value=0;form2_"&rsvisual("id")&".fromuser.value='';openScript('about:blank',500,400);form2_"&rsvisual("id")&".submit()"">����</font></b>"
    			end if
      	end if
      	response.write "��"
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
				response.write "</td>"
		    if curcount mod 2=0 then
					response.write "</tr>"
					response.write "</table>"
					response.write "</td>"
					response.write "</tr>"
				end if
				rsvisual.movenext
			loop
		end if
		curcount=curcount+1
		for i=curcount to 10
			if i mod 2=1 then
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody1 align=center>"
				response.write "<table width=560 border=0 cellspacing=0 cellpadding=0>"
				response.write "<tr>"
			else
				response.write "<td width=14>&nbsp;</td>"
			end if
			response.write "<td width=282>"
			response.write "<table width=282 border=0 cellspacing=2 cellpadding=0>"
			response.write "<tr>"
			response.write "<td rowspan=5 height=84 width=84 class=tablebody2>&nbsp;</td>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "<td class=tablebody2 height=20>&nbsp;</td>"
			response.write "</tr>"
			response.write "</table>"
			response.write "</td>"
			if i mod 2=0 then
				response.write "</tr>"
				response.write "</table>"
				response.write "</td>"
				response.write "</tr>"
			end if
		next
		rsvisual.close
		response.write "</table>"
		response.write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
		response.write "<tr height=25 valign=middle>"
		response.write "<td valign=middle class=tablebody1>&nbsp;&nbsp;ҳ�Σ�<b>"&curPage&"</b> / <b>"&pagecount&"</b> ҳ ÿҳ��<b>10</b> ���ƣ�<b>"&totalrec&"</b></td>"
		response.write "<td valign=middle class=tablebody1 align=right>��ҳ��"
		call DispPageNum(curpage,pagecount,"?shopid="&shopid&"&sex="&sex&"&page=","&search_str="&request("search_str")&"&selgender="&request("selgender")&"&selprice_l="&request("selprice_l")&"&selprice_h="&request("selprice_h")&"&selperiod_l="&request("selperiod_l")&"&selperiod_h="&request("selperiod_h")&"&selflag="&request("selflag")&"&sortobj="&request("sortobj")&"&sorttype="&request("sorttype")&"#VisualTop")
		response.write "&nbsp;&nbsp;</td>"
		response.write "</tr>"
		response.write "</table>"
	elseif CurType=98 then
		response.write "<br>"
		response.write "<table border=0 cellspacing=1 cellpadding=0 bgcolor="&Forum_Body(27)&" width=95% >"
		response.write "<tr height=25 valign=middle>"
		response.write "<th valign=middle>ʹ�ð���</th>"
		response.write "</tr>"
		response.write "<tr height=620>"
		response.write "<td width=100% class=tablebody2 valign=top><br><blockquote style=""line-height:20px"">"
		if shopid=198 then
			response.write "������ӭ��ʹ��<font color=blue>"&forum_info(0)&"</font>��<b>�������ϵͳ</b>��<br><br>"
			response.write "������ϵͳ��Ϊ<a href=?shopid=298><font color=blue><b>������Ʒ</b></font></a>��<a href=?shopid=398><font color=blue><b>�����̵�</b></font></a>��<a href=?shopid=498><font color=blue><b>�����г�</b></font></a>��<a href=?shopid=598><font color=blue><b>�����</b></font></a>�Ĳ��֡�<br><br>"
			response.write "���������ĸ��������Դ���¼�͸���״̬�£������Կ�����<b>�鿴�����</b>�������ӣ�������е���Ʒ����ɾ������գ����͵���Ʒ���Ե���������޸���ͬ���ԣ�������Ϊ������Ʒ��û������󣬵����<B>����</b>������֧��������õ����ͳ��������Ʒ��������������õ���Ʒ����ʱ�Ѿ��ŵ����Ĵ�����У���������Ʒ�Ѿ����Դ�ʱ���������ϣ����ʱϵͳҲ���Զ���������һ���µ�����<br><br>"
			response.write "�����ڡ�<b>�鿴�����</b>�����������ĸ����ӿ��Թ����˽��������������͡������Լ�������Ʒ�ļ�¼��<br><br>"
		elseif shopid=298 then
			response.write "<a href=""?shopid=100""><font color="&forum_body(8)&"><b>������Ʒ</b></font></a>��<br><br>"
			response.write "�������Բ鿴�Լ��Ѿ�ӵ�е���Ʒ����������ʹ�õķ�װ����ñ�ͱ����ȸ�����Ʒ�������Ѿ�ӵ�е���Ʒ���Լ�����Լ����ڳ��۵���Ʒ�ȡ�"
			response.write "������Ʒ���裺<br><br>"
			response.write "����1��<a href=""?shopid=100""><font color=blue><B>�ҵ��³�</b></font></a>�����ڴ����װ�����¡���ȹ��ñ�ӵ���Ʒ��<br><br>"
			response.write "����2��<a href=""?shopid=200""><font color=blue><B>�ҵĻ�ױ��</b></font></a>�����ڴ��ͷ�����͡����Ρ���ü�����͡��ڱǺͱ������Ʒ��<br><br>"
			response.write "����3��<a href=""?shopid=300""><font color=blue><B>�ҵ����κ�</b></font></a>�����ڴ������Ʒ��ͷ��Ʒ����Ʒ���۾��Ͱ�������Ʒ��<br><br>"
			response.write "����4��<a href=""?shopid=400""><font color=blue><B>�ҵ�������</b></font></a>�����ڴ�ű�����Ч����Q��Q�á���������Լ�������Ʒ�ȡ�<br><br>"
			response.write "���������������ƷֻҪû�й��ڣ���������ʱ���ͼƬ�Դ���<B>����</b>�����ɻ��ϣ�Ҳ���ԡ�<b>����</b>�������ĺ��ѣ�����ʱ����һ��д�ϼ���ף���Ļ��һ����Ʒ�������ڣ����Ե㡰<B>����</b>�������޸����޸�ʱ��Ҫѡ��������һ������Ч�ڣ���������Զ��Ч�����ڲ���Ҫ����Ʒ�����Էŵ������г���<B>����</b>��������ǰ��Ҫָ���ۼ۲�֧�������ѡ�<br><br>"
			response.write "����5��<a href=""?shopid=500""><font color=blue><B>�ҵĻ���</b></font></a>���ڴˣ������Կ������ڶ����г����۵�������Ʒ����Щ��Ʒ��һ�����Ե��ͼƬ�Դ���<B>����</b>����ͬʱ��Ҳ���ԡ�<B>����</b>���͡�<B>����</b>����Щ�Ѿ����۵���Ʒ������ʱϵͳ���Զ�ȡ�����ĳ��۲��˻����������ѣ�������������������������Ʒ�����Ե����<B>ͣ��</b>���������ѻ������˻�������<br><br>"
			response.write "����6��<a href=""?shopid=600""><font color=blue><B>�ҵ�����¨</b></font></a>���������Ʒ�����Ѿ����ڵģ�������޷��ڴ��Դ���Щ��Ʒ����ֻ�С�<b>����</b>����ſ��Լ���ʹ�ã��������Ʒһ�����ԡ�<B>����</b>���͡�<B>����</b>����<br><br>"
		elseif shopid=398 then
			response.write "<a href=""?shopid=101""><font color="&forum_body(8)&"><b>�����̵�</b></font></a>��<br><br>"
			response.write "��������"
			sqlvisual="select seqno,typename from visual_type where visible order by orders"
			rsvisual.open sqlvisual,conn,1,1
			set rs=server.createobject("adodb.recordset")
			TypeCount=0
			do while not rsvisual.eof
				sqlvisual="select top 1 seqno from visual_subtype where typeseq="&rsvisual("seqno")&" and visible order by orders"
				rs.open sqlvisual,conn,1,1
				if not rs.eof then
					if TypeCount>0 then
						rsvisual.movenext
						if rsvisual.eof then
							response.write "��"
						else
							response.write "��"
						end if
						rsvisual.moveprevious
					end if
					response.write "<a href=""?shopid="&rs("seqno")&right("00"&rsvisual("seqno"),2)&"""><font color=blue><B>"&rsvisual("TypeName")&"</b></font></a>"
					TypeCount=TypeCount+1
				end if
				rs.close
				rsvisual.movenext
			loop
			set rs=nothing
			rsvisual.close
			response.write "����"&TypeCount&"��"
			response.write "��̨��������������ѡ�����������Ʒ�������<B>�Դ�</b>���󱣴��ֱ�ӵ����<B>����</b>���������Ĺ���������㡰<B>����</b>������ƷҲ����빺����У�����������Դ��<font color="&forum_body(8)&"><b>"&CartLength&"</b></font>����Ʒ�������˾ͱ��롰<B>����</b>������ܼ���ѡ����"
			response.write "��Щ��Ʒֻ���ʺ�ĳЩ��Աʹ�ã��빺��ʱע�⡣"
			response.write "<font color="&forum_body(8)&"><b>VIP��Ա</b></font>�����򵽸�������˵���Ʒ��"
			response.write "�����Ʒ�Ѿ��ϻ������Ͳ��������ˡ�<br><br>"
			response.write "����<a href=""?shopid=199""><font color=blue><B>�Ƽ���Ʒ</b></font></a>��<font color=blue>"&forum_info(0)&"</font>�ṩ�����ļ򵥲����򵼣������<a href=""?shopid=199""><font color=blue><B>�����ϼ�</b></font></a>�г��������̵�����½���������Ʒ��<a href=""?shopid=299""><font color=blue><B>������Ʒ</b></font></a>���г�Ŀǰ������Ա��ӵ����������Ʒ��<a href=""?shopid=399""><font color=blue><B>VIPר��</b></font></a>�п��Կ����̵������ж�<font color="&forum_body(8)&"><b>VIP��Ա</b></font>ר������Ʒ�����������<font color="&forum_body(8)&"><b>VIP��Ա</b></font>����Ͽ�<a href=z_vip.asp target=_blank><font color=blue><b>����</b></font></a>�ɣ���ȻҲ���Ե�<a href=""?shopid=196""><font color=blue><B>�����г�</b></font></a>תת��˵�����������ع���ƷӴ��<br><br>"
		elseif shopid=498 then
			response.write "<a href=""?shopid=196""><font color="&forum_body(8)&"><B>�����г�</b></font></a>��<br><br>"
			response.write "�������ڼ۸���ˣ���Ʒ�����廨���ţ�������Է��������ϲ������Ʒ�޷��Դ���ֻ��ѡ��������<B>����</b>����������<B>����</b>���������г�û��Ϊ��׼���������ѡ������Ʒ���������ֽ��ס�"
			response.write "��������󣬳�����Ҳ�������õ���֧���Ŀ��<br><br>"
		elseif shopid=598 then
			response.write "<a href=""?shopid=197""><font color="&forum_body(8)&"><B>�����</b></font></a>��<br><br>"
			response.write "����1��<b>����д��</b>��<br>"
			response.write "����д��ǰ����Ҫ�Ȼ����Լ�׼��д����·��������·�ֻ��<b>�Դ�</b>������Ҫ<b>����</b>��Ȼ���������±ߵ�<b>д��</b>���ɽ������д�����ơ�"
			response.write "���ʱ���Էֱ�ѡ����������ƣ��Լ�д������ƺ�д���ʱ������ֺ�ͼ������ƶ������á���������ת������Ȳ����������Ʊ����д�����Ƭ�������ɡ�<br><br>"
			response.write "����2��<b>���˺�Ӱ</b>��<br>"
			response.write "���������ṩ��� <b><font color="&forum_body(8)&">"&conn.execute("select top 1 PhotoUserCount from Visual_Config")(0)&"��</font></b> �ļ����Ӱ����Ӱǰ��������Ҫ�Ȼ����Լ��μӺ�Ӱ���·��������·�ֻ��<b>�Դ�</b>������Ҫ<b>����</b>��Ȼ�����<a href=?shopid=197><font color=blue><b>����������</b></font></a>ҳ�棬����ȷ����Ӱ�������Լ��Բ����ߵ������ԣ�ͬʱ����Ҫָ����Ƭ�ı����Լ���Ƭ�Ƿ񹫿�����������ȷ����ϵͳ�����������б�����Ĳ����߷��Ͷ���Ϣ��"
			response.write "��������<a href=?shopid=297><font color=blue><b>���͵�����</b></font></a>�п�������������������δ�����в�����ȫ��ȷ�ϵ�����<br><br>"
			response.write "�������������յ��������Ϣ����<a href=?shopid=397><font color=blue><b>�յ�������</b></font></a>���ҵ��Լ���δȷ�ϵ��������ȷ�ϣ�ͬ����������Ҳ������ȷ��ǰ�����Լ�������·���ͬ�������·�ʱֻ��<b>�Դ�</b>����<b>����</b>��ȷ��ʱ���Կ�����Ƭ�г��ֵ���������Ѿ�ȷ�ϵĻ�Ա��"
			response.write "�����л�Աȫ��ȷ��ǰ�����˷������⣬���л�Ա�������ٴθ�����װ�����µ���������ȷ������"
			response.write "�����в�����ȫ��ȷ�Ϻ�ϵͳ��֪ͨ��������Ƭ�Ѿ���������ȷ�ϣ����Կ�ʼ����ˡ�<br><br>"
			response.write "���������߿��Խ���<a href=?shopid=497><font color=blue><b>ȷ�ϵ�����</b></font></a>�в鿴��Щ�����Ѿ������в�����ȷ���ˣ��ȶ����Զ���Щȷ�Ϻ����Ƭ������ƣ����ʱ���Էֱ�ѡ����������ƣ��Լ���Ӱ�����ƺͺ�Ӱ��ʱ������ֺ�ͼ������ƶ������á���������ת������Ȳ����������Ʊ����ϵͳ��֪ͨ���в����߶�������Ʒ��������"
			response.write "�����в�����ȫ��ͬ��ǰ������Ȼ���Լ����޸���Ƭ����ƣ�һ�����в�����ȫ��ͬ����������ƺ�ϵͳ��֪ͨ��Ƭ�е�ÿ���ˣ���Ƭ�Ѿ���ɡ�<br><br>"
			response.write "����3��<b>������</b>��<br>"
			response.write "����<a href=?shopid=597><font color=blue><b>�ҵ�������</b></font></a>�б����������������к�Ӱ������Ҳ������Щ���ܵĺ�Ӱ�������������ѡ��ɾ����Щ������Ҫ����Ƭ��ɾ���������Լ���������Щ��Ƭ�����ڳ��֣����ǣ�������в���������Ȼ���˱�����������Ƭ����ô��<a href=?shopid=697><font color=blue><b>��ҵ���</b></font></a>����Ȼ�����������Ƭ��ֻ�����в����߶�ɾ���ˣ���ô������Ƭ�Ż᳹����ʧ��"
			response.write "�����л���һ��<b>��ӡ</b>���ܣ������԰�������������ɵ�PNG�ļ�����������ʱ������������ļ���<br><br>"
			response.write "����<a href=?shopid=697><font color=blue><b>��ҵ�������</b></font></a>�б���������л�Ա�����ĺ�Ӱ��<br><br>"
		end if
		response.write "</blockquote><br></td>"
		response.write "</tr>"
		response.write "</table>"
		response.write "<br>"
	elseif CurType=97 then%>
		<!--#include file="z_visual_js.asp"-->
		<%response.write "<br>"
		if shopid=197 then
			call DispPhoto(0)
		elseif shopid=297 then
			if isnull(request("photoid")) then
				call dispphotolist(1)
			elseif not isInteger(request("photoid")) then
				call dispphotolist(1)
			elseif conn.execute("select id from visual_photo where id="&checkstr(trim(request("photoid")))&" and fromuser='"&membername&"' and not confirmed").eof then
				call dispphotolist(1)
			else
				call dispphoto(1)
			end if
		elseif shopid=397 then
			if isnull(request("photoid")) then
				call dispphotolist(2)
			elseif not isInteger(request("photoid")) then
				call dispphotolist(2)
			elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and not p.confirmed").eof then
				call dispphotolist(2)
			else
				call dispphoto(2)
			end if
		elseif shopid=497 then
			if isnull(request("photoid")) then
				call dispphotolist(3)
			elseif not isInteger(request("photoid")) then
				call dispphotolist(3)
			elseif conn.execute("select u.photo_id from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.photo_id="&checkstr(trim(request("photoid")))&" and u.username='"&membername&"' and p.confirmed and not p.finished").eof then
				call dispphotolist(3)
			else
				call dispphoto(3)
			end if
		else
			response.write "<table border=0 cellspacing=1 cellpadding=0 bgcolor="&Forum_Body(27)&" width=590 >"
			response.write "<tr height=25 valign=middle>"
			response.write "<th colspan=3>"&forum_info(0)&"����� - "
			if shopid=597 then
				sqlvisual="select p.* from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.username='"&membername&"' and p.finished and not u.deleted order by p.FinishDate desc"
				response.write "�ҵ������ಾ"
			else
				sqlvisual="select * from visual_photo where status=1 and finished order by FinishDate desc"
				response.write "��ҵ������ಾ"
			end if
			response.write "</th>"
			response.write "</tr>"
			rsvisual.open sqlvisual,conn,1,1
			totalrec=rsvisual.recordcount
			if totalrec mod 4=0 then
				pagecount=totalrec \ 4
			else
				pagecount=totalrec \ 4+1
			end if
			curcount=0
			dim flag
			flag=false
			if not rsvisual.eof then
				flag=true
				if curpage > pagecount then curpage = pagecount
				if curpage < 1 then curpage=1
				Rsvisual.PageSize=4
				Rsvisual.AbsolutePage=CurPage
				dim rsPhotoUser
				set rsPhotoUser=server.createobject("ADODB.recordset")
				dim usercount,uservisualsplit(),usernamesplit(),PhotoDate
				dim PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection
				dim DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection
				dim NameLeft(),NameTop(),NameFont(),NameSize(),NameBold(),NameItalic(),NameColor(),NameDirection()
				dim OuterLeft(),OuterTop(),OuterWidth(),OuterHeight()
				dim InnerLeft(),InnerTop(),InnerWidth(),InnerHeight()
				dim LayerNo(),Direction()
				dim UserNameStr
				response.write "<tr valign=middle>"
				response.write "<td class=tablebody1 align=center>"
				response.write "<table width=590 border=0 cellspacing=0 cellpadding=0>"
				do while not rsvisual.eof and curcount<4
					curcount=curcount+1
					sqlvisual="select * from visual_photouser where photo_id="&rsvisual("ID")
					rsPhotoUser.open sqlvisual,conn,1,1
					usercount=0
					UserNameStr=""
					do while not rsPhotoUser.eof
						redim preserve uservisualsplit(usercount)
						uservisualsplit(usercount)=rsPhotoUser("uservisual")
						redim preserve usernamesplit(usercount)
						usernamesplit(usercount)=rsPhotoUser("username")
						redim preserve LayerNo(UserCount)
						redim preserve OuterLeft(UserCount)
						redim preserve OuterTop(UserCount)
						redim preserve OuterWidth(UserCount)
						redim preserve OuterHeight(UserCount)
						redim preserve InnerLeft(UserCount)
						redim preserve InnerTop(UserCount)
						redim preserve InnerWidth(UserCount)
						redim preserve InnerHeight(UserCount)
						redim preserve Direction(UserCount)
						redim preserve NameLeft(UserCount)
						redim preserve NameTop(UserCount)
						redim preserve NameFont(UserCount)
						redim preserve NameSize(UserCount)
						redim preserve NameBold(UserCount)
						redim preserve NameItalic(UserCount)
						redim preserve NameColor(UserCount)
						redim preserve NameDirection(UserCount)
						LayerNo(UserCount)=rsPhotoUser("LayerNo")
						OuterLeft(UserCount)=rsPhotoUser("OuterLeft")
						OuterTop(UserCount)=rsPhotoUser("OuterTop")
						OuterWidth(UserCount)=rsPhotoUser("OuterWidth")
						OuterHeight(UserCount)=rsPhotoUser("OuterHeight")
						InnerLeft(UserCount)=rsPhotoUser("InnerLeft")
						InnerTop(UserCount)=rsPhotoUser("InnerTop")
						InnerWidth(UserCount)=rsPhotoUser("InnerWidth")
						InnerHeight(UserCount)=rsPhotoUser("InnerHeight")
						Direction(UserCount)=rsPhotoUser("Direction")
						NameLeft(UserCount)=rsPhotoUser("NameLeft")
						NameTop(UserCount)=rsPhotoUser("NameTop")
						NameFont(UserCount)=rsPhotoUser("NameFont")
						NameSize(UserCount)=rsPhotoUser("NameSize")
						NameBold(UserCount)=rsPhotoUser("NameBold")
						NameItalic(UserCount)=rsPhotoUser("NameItalic")
						NameColor(UserCount)=rsPhotoUser("NameColor")
						NameDirection(UserCount)=rsPhotoUser("NameDirection")
						usercount=usercount+1
						if usercount<=2 then UserNameStr=UserNameStr&rsPhotoUser("username")
						rsPhotoUser.movenext
						if not rsPhotoUser.eof and usercount<2 then UserNameStr=UserNameStr&","
					loop
					if usercount>2 then UserNameStr=UserNameStr&" ��"&UserCount&"��"
					rsPhotoUser.close
					sqlvisual="select * from visual_Accouterment where photo_id="&rsvisual("ID")
					rsPhotoUser.open sqlvisual,conn,1,1
					redim preserve uservisualsplit(usercount+5)
					redim preserve OuterLeft(UserCount+5)
					redim preserve OuterTop(UserCount+5)
					redim preserve OuterWidth(UserCount+5)
					redim preserve OuterHeight(UserCount+5)
					redim preserve InnerLeft(UserCount+5)
					redim preserve InnerTop(UserCount+5)
					redim preserve InnerWidth(UserCount+5)
					redim preserve InnerHeight(UserCount+5)
					redim preserve Direction(UserCount+5)
					for i=0 to 5
						uservisualsplit(usercount+i)=""
						OuterLeft(UserCount+i)=0
						OuterTop(UserCount+i)=0
						OuterWidth(UserCount+i)=140
						OuterHeight(UserCount+i)=226
						InnerLeft(UserCount+i)=0
						InnerTop(UserCount+i)=0
						InnerWidth(UserCount+i)=140
						InnerHeight(UserCount+i)=226
						Direction(UserCount+i)=0
					next
					do while not rsPhotoUser.eof
						i=rsPhotoUser("SeqNo")
						uservisualsplit(usercount+i)=rsPhotoUser("ItemPicPath")
						OuterLeft(UserCount+i)=rsPhotoUser("OuterLeft")
						OuterTop(UserCount+i)=rsPhotoUser("OuterTop")
						OuterWidth(UserCount+i)=rsPhotoUser("OuterWidth")
						OuterHeight(UserCount+i)=rsPhotoUser("OuterHeight")
						InnerLeft(UserCount+i)=rsPhotoUser("InnerLeft")
						InnerTop(UserCount+i)=rsPhotoUser("InnerTop")
						InnerWidth(UserCount+i)=rsPhotoUser("InnerWidth")
						InnerHeight(UserCount+i)=rsPhotoUser("InnerHeight")
						Direction(UserCount+i)=rsPhotoUser("Direction")
						rsPhotoUser.movenext
					loop
					rsPhotoUser.close
					if curcount mod 2=1 then
						response.write "<tr>"
					else
						response.write "<td width=14>&nbsp;</td>"
					end if
					response.write "<td width=290>"
					response.write "<table width=290 border=0 cellspacing=2 cellpadding=0>"
					response.write "<tr valign=middle height=226>"
					response.write "<td class=tablebody2 align=center colspan=2 width=100% >"
					call ShowVisualPhoto(rsvisual("Width"),rsvisual("Height"),rsvisual("Background"),rsvisual("BackBody"),rsvisual("ForeBody"),rsvisual("Foreground"),rsvisual("Name"),rsvisual("PhotoNameLeft"),rsvisual("PhotoNameTop"),rsvisual("PhotoNameFont"),rsvisual("PhotoNameSize"),rsvisual("PhotoNameBold"),rsvisual("PhotoNameItalic"),rsvisual("PhotoNameColor"),rsvisual("PhotoNameDirection"),rsvisual("AddDate"),rsvisual("DateLeft"),rsvisual("DateTop"),rsvisual("DateFont"),rsvisual("DateSize"),rsvisual("DateBold"),rsvisual("DateItalic"),rsvisual("DateColor"),rsvisual("DateDirection"),UserCount,UserVisualSplit,LayerNo,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction,UserNameSplit,NameLeft,NameTop,NameFont,NameSize,NameBold,NameItalic,NameColor,NameDirection,False,rsVisual("isUpload"),(rsVisual("Child")=1))
					response.write "</td>"
					response.write "</tr>"
					response.write "<tr valign=middle height=22>"
					response.write "<td class=tablebody2 width=80 align=right>"
					response.write "��Ƭ����&nbsp;"
					response.write "</td>"
					response.write "<td class=tablebody2 align=left>"
					response.write "&nbsp;"&rsvisual("name")
					response.write "</td>"
					response.write "</tr>"
					response.write "<tr valign=middle height=22>"
					response.write "<td class=tablebody2 width=80 align=right>"
					response.write "����ʱ��&nbsp;"
					response.write "</td>"
					response.write "<td class=tablebody2 align=left>"
					response.write "&nbsp;"&formatdatetime(rsvisual("AddDate"),2)
					response.write "</td>"
					response.write "</tr>"
					response.write "<tr valign=middle height=22>"
					response.write "<td class=tablebody2 width=80 align=right>"
					response.write "������&nbsp;"
					response.write "</td>"
					response.write "<td class=tablebody2 align=left>"
					response.write "&nbsp;"&UserNameStr
					response.write "</td>"
					response.write "</tr>"
					response.write "<tr valign=middle height=22>"
					response.write "<td class=tablebody2 width=80 align=right>"
					response.write "��ƬURL&nbsp;"
					response.write "</td>"
					response.write "<td class=tablebody2 align=left>"
					if not isnull(rsvisual("URL")) then response.write "&nbsp;<a href="&rsvisual("URL")&" target=_blank><font color=blue><b>�鿴��������Ƭ</b></font></a>"
					response.write "</td>"
					response.write "</tr>"
					if shopid=597 or master then
						response.write "<tr valign=middle height=22>"
						response.write "<td class=tablebody2 width=80 align=right>"
						response.write "����&nbsp;"
						response.write "</td>"
						response.write "<td class=tablebody2 align=left>"
						response.write "&nbsp;<a href=z_visual_photo.asp?action=delete&photoid="&rsvisual("ID")&" onclick=""javascript:if(confirm('��ȷ��ɾ��������Ƭ��?')){return true;} return false;"">ɾ��</a>&nbsp;<a href=z_visual_photo.asp?action=recreate&photoid="&rsvisual("ID")&">��ӡ</a>"
						if lcase(trim(rsvisual("fromUser")))=lcase(trim(membername)) or master then
							if rsvisual("status")=1 then
								response.write "&nbsp;<a href=z_visual_photo.asp?action=privatephoto&photoid="&rsvisual("ID")&">����</a>"
							else
								response.write "&nbsp;<a href=z_visual_photo.asp?action=publicphoto&photoid="&rsvisual("ID")&">����</a>"
							end if
						end if
						response.write "</td>"
						response.write "</tr>"
					end if
					response.write "</table>"
					response.write "</td>"
			    if curcount mod 2=0 then
						response.write "</tr>"
					end if
					rsvisual.movenext
				loop
				set rsPhotoUser=nothing
				rsvisual.close
			end if
			curcount=curcount+1
			for i=curcount to 4
				if i mod 2=1 then
					response.write "<tr>"
				else
					response.write "<td width=14>&nbsp;</td>"
				end if
				response.write "<td width=290>"
				response.write "<table width=290 border=0 cellspacing=2 cellpadding=0>"
				response.write "<tr valign=middle height=226>"
				response.write "<td class=tablebody2 align=center colspan=2>"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle height=22>"
				response.write "<td class=tablebody2 width=80 align=right>"
				response.write "</td>"
				response.write "<td class=tablebody2 align=left>"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle height=22>"
				response.write "<td class=tablebody2 width=80 align=right>"
				response.write "</td>"
				response.write "<td class=tablebody2 align=left>"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle height=22>"
				response.write "<td class=tablebody2 width=80 align=right>"
				response.write "</td>"
				response.write "<td class=tablebody2 align=left>"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr valign=middle height=22>"
				response.write "<td class=tablebody2 width=80 align=right>"
				response.write "</td>"
				response.write "<td class=tablebody2 align=left>"
				response.write "</td>"
				response.write "</tr>"
				if shopid=597 or master then
					response.write "<tr valign=middle height=22>"
					response.write "<td class=tablebody2 width=80 align=right>"
					response.write "</td>"
					response.write "<td class=tablebody2 align=left>"
					response.write "</td>"
					response.write "</tr>"
				end if
				response.write "</table>"
				response.write "</td>"
		    if i mod 2=0 then
					response.write "</tr>"
				end if
			next
			response.write "</table>"
			if flag then
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
			end if
			response.write "<table width=590 border=0 cellspacing=0 cellpadding=0>"
			response.write "<tr height=25 valign=middle>"
			response.write "<td valign=middle class=tablebody1>&nbsp;&nbsp;ҳ�Σ�<b>"&curPage&"</b> / <b>"&pagecount&"</b> ҳ ÿҳ��<b>4</b> ���ƣ�<b>"&totalrec&"</b></td>"
			response.write "<td valign=middle class=tablebody1 align=right>��ҳ��"
			call DispPageNum(curpage,pagecount,"?shopid="&shopid&"&sex="&sex&"&page=","#VisualTop")
			response.write "&nbsp;&nbsp;</td>"
			response.write "</tr>"
			response.write "</table>"
		end if
		response.write "<br>"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "</table>"
	set rsvisual=nothing
end sub

sub DispPhotoList(ListMethod)
	dim totalrec,pagecount,curcount
	response.write "<table border=0 cellspacing=1 cellpadding=0 bgcolor="&Forum_Body(27)&" width=95% >"
	response.write "<tr height=25 valign=middle>"
	response.write "<th colspan=6>"&forum_info(0)&"����� - "
	select case ListMethod
	case 1
		response.write "����������"
	case 2
		response.write "�յ�������"
	case 3
		response.write "ȷ�ϵ�����"
	end select
	response.write "</th>"
	response.write "</tr>"
	response.write "<tr valign=middle height=50>"
	response.write "<td class=tablebody2 colspan=6 align=center>"
	select case ListMethod
	case 1
		response.write "����������δ�����в�����ȷ�ϣ�����δ��ȷ�ϵ����󽫱�ϵͳ�Զ�ɾ����"
	case 2
		response.write "������ϣ��������Ӱ��������Щ�������ȴ��������������ߵ�ȷ�Ϸ��ɣ��������ƺ�ȷ������"
	case 3
		response.write "�����������������Ѿ������в�����ȷ���ˣ����������޸���Ƭ����ơ�"
	end select
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25 valign=middle>"
	response.write "<th width=""*"">����</th>"
	response.write "<th width=""15%"">������</th>"
	response.write "<th width=""10%"">״̬</th>"
	response.write "<th width=""12%"">����ʱ��</th>"
	response.write "<th width=""12%"">����ʱ��</th>"
	response.write "<th width=""15%"">������</th>"
	response.write "</tr>"
	dim rsphotouser
	set rsphotouser=server.createobject("adodb.recordset")
	select case ListMethod
	case 1
		sqlvisual="select id,name,status,adddate,enddate,fromuser from visual_photo where fromuser='"&membername&"' and not confirmed"
	case 2
		sqlvisual="select p.id,p.name,p.status,p.adddate,p.enddate,p.fromuser from visual_photouser u inner join visual_photo p on u.photo_id=p.id where p.fromuser<>'"&membername&"' and not p.confirmed and u.username='"&membername&"'"
	case 3
		sqlvisual="select p.id,p.name,p.status,p.adddate,p.enddate,p.fromuser from visual_photouser u inner join visual_photo p on u.photo_id=p.id where u.username='"&membername&"' and p.confirmed and not p.finished"
	end select
	rsvisual.open sqlvisual,conn,1,1
	totalrec=rsvisual.recordcount
	if totalrec mod 10=0 then
		pagecount=totalrec \ 10
	else
		pagecount=totalrec \ 10+1
	end if
	curcount=0
	if rsvisual.eof or rsvisual.bof then
		response.write "<tr valign=middle height=30><td class=tablebody2 colspan=6 align=center>"
		select case ListMethod
		case 1
			response.write "��δ�����κκ�Ӱ����������е������ѵõ�ȷ���ˡ�"
		case 2
			response.write "û���յ��κκ�Ӱ����������е������ѱ���ȷ���ˡ�"
		case 3
			response.write "û���κ�ȷ����ϲ��ȴ���ɵĺ�Ӱ��������д�档"
		end select
		response.write "</td></tr>"
	else
		if curpage > pagecount then curpage = pagecount
		if curpage < 1 then curpage=1
		Rsvisual.PageSize=10
		Rsvisual.AbsolutePage=CurPage
		do while not rsvisual.eof and curcount<10
			curcount=curcount+1
			response.write "<tr height=25 valign=middle>"
			response.write "<td class=tablebody2>"
			response.write "&nbsp;<a href=""z_Visual.asp?shopid="
			select case ListMethod
			case 1
				response.write "297"
			case 2
				response.write "397"
			case 3
				response.write "497"
			end select
			response.write "&photoid="&rsvisual(0)&""">"&rsvisual(1)&"</a>"
			response.write "</td>"
			response.write "<td class=tablebody2 align=center>"
			sqlvisual="select username,confirmed,confirmtime,finished from visual_photouser where photo_id="&rsvisual(0)&" and username='"&rsvisual(5)&"'"
			rsphotouser.open sqlvisual,conn,1,1
			response.write "<a href=""dispuser.asp?name="&rsvisual(5)&""" target=_blank title="""
			if ListMethod=3 then
				if rsphotouser(3) then
					response.write "�Ѿ������Ƭ�����"
				else
					response.write "��δ�����Ƭ�����"
				end if
			else
				if rsphotouser(1) then
					response.write "����"&formatdatetime(rsphotouser(2),2)&"ȷ��"
				else
					response.write "��δȷ��"
				end if
			end if
			response.write """>"
			if ListMethod=3 then
				if not rsphotouser(3) then
					response.write "<font color="&forum_body(8)&">"
				end if
			else
				if not rsphotouser(1) then
					response.write "<font color="&forum_body(8)&">"
				end if
			end if
			response.write rsvisual(5)
			if ListMethod=3 then
				if not rsphotouser(3) then
					response.write "</font>"
				end if
			else
				if not rsphotouser(1) then
					response.write "</font>"
				end if
			end if
			response.write "</a>"
			rsphotouser.close
			response.write "</td>"
			response.write "<td class=tablebody2 align=center>"
			if rsvisual(2)=1 then
				response.write "����"
			else
				response.write "����"
			end if
			response.write "</td>"
			response.write "<td class=tablebody2 align=center>"
			response.write formatdatetime(rsvisual(3),2)
			response.write "</td>"
			response.write "<td class=tablebody2 align=center>"
			response.write formatdatetime(rsvisual(4),2)
			response.write "</td>"
			response.write "<td class=tablebody2 align=center>"
			sqlvisual="select username,confirmed,confirmtime,finished from visual_photouser where photo_id="&rsvisual(0)&" and username<>'"&rsvisual(5)&"'"
			rsphotouser.open sqlvisual,conn,1,1
			if not rsphotouser.eof then
				do while not rsphotouser.eof
					response.write "<a href=""dispuser.asp?name="&rsphotouser(0)&""" target=_blank title="""
					if ListMethod=3 then
						if rsphotouser(3) then
							response.write "�Ѿ�ͬ����Ƭ�����"
						else
							response.write "��δͬ����Ƭ�����"
						end if
					else
						if rsphotouser(1) then
							response.write "����"&formatdatetime(rsphotouser(2),2)&"ȷ��"
						else
							response.write "��δȷ��"
						end if
					end if
					response.write """>"
					if ListMethod=3 then
						if not rsphotouser(3) then
							response.write "<font color="&forum_body(8)&">"
						end if
					else
						if not rsphotouser(1) then
							response.write "<font color="&forum_body(8)&">"
						end if
					end if
					response.write rsphotouser(0)
					if ListMethod=3 then
						if not rsphotouser(3) then
							response.write "</font>"
						end if
					else
						if not rsphotouser(1) then
							response.write "</font>"
						end if
					end if
					response.write "</a>"
					rsphotouser.movenext
					if not rsphotouser.eof then response.write "<br>"
				loop
			else
				response.write "<font color=gray>����д��</font>"
			end if
			rsphotouser.close
			response.write "</td>"
			response.write "</tr>"
			rsvisual.movenext
		loop
		rsvisual.close
	end if
	response.write "</table>"
	response.write "<table width=95% border=0 cellspacing=0 cellpadding=0>"
	response.write "<tr height=25 valign=middle>"
	response.write "<td valign=middle class=tablebody1>&nbsp;&nbsp;ҳ�Σ�<b>"&curPage&"</b> / <b>"&pagecount&"</b> ҳ ÿҳ��<b>10</b> ���ƣ�<b>"&totalrec&"</b></td>"
	response.write "<td valign=middle class=tablebody1 align=right>��ҳ��"
	call DispPageNum(curpage,pagecount,"?shopid="&shopid&"&sex="&sex&"&page=","#VisualTop")
	response.write "&nbsp;&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"
end sub

sub DispPhoto(ListMethod)
	dim photoid,subtype,photoname,photomsg,fromuser,photostatus,bgwidth,bgheight,photobg,photoBodyBack,photoBodyFore,photoFg
	dim usercount,uservisualsplit(),usernamesplit(),LeftPos,StepPos,FinishDate,PhotoDate
	dim PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection
	dim DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection
	dim NameLeft(),NameTop(),NameFont(),NameSize(),NameBold(),NameItalic(),NameColor(),NameDirection()
	dim OuterLeft(),OuterTop(),OuterWidth(),OuterHeight()
	dim InnerLeft(),InnerTop(),InnerWidth(),InnerHeight()
	dim LayerNo(),Direction()
	dim hasFinished,isUpload

	FinishDate=Null
	hasFinished=False
	select case ListMethod
	case 0
		response.write "<form action=""z_visual_photo.asp"" method=post name=photo_request>"
		response.write "<input type=hidden name=""action"" value=""send"">"
		if request.cookies("myshow_"&userid)="" then
			response.write "<input type=hidden name=""fromvisual"" value="""&v_myVisual&""">"
		else
			response.write "<input type=hidden name=""fromvisual"" value="""&request.cookies("myshow_"&userid)&""">"
		end if
		response.write "<input type=hidden name=""bgwidth"" value=""280"">"
		response.write "<input type=hidden name=""bgheight"" value=""226"">"
		subtype="����������"
		photoname=""
		photomsg=""
		usercount=1
		redim uservisualsplit(0)
		redim usernamesplit(0)
		if request.cookies("myshow_"&userid)="" then
			uservisualsplit(0)=v_myvisual
		else
			uservisualsplit(0)=request.cookies("myshow_"&userid)
		end if
		usernamesplit(0)=membername
		bgwidth=280
		bgheight=226
		photobg=""
		photoBodyBack=""
		photoBodyFore=""
		photoFg=""
		PhotoDate=now()
		isUpload=false
	case 1
		photoid=checkstr(trim(request("photoid")))
		sqlvisual="select * from visual_photo where id="&photoid
		rsvisual.open sqlvisual,conn,1,1
		subtype="����������"
		photoname=rsvisual("name")
		photomsg=rsvisual("content")
		fromuser=rsvisual("fromuser")
		photostatus=rsvisual("status")
		bgwidth=rsvisual("width")
		bgheight=rsvisual("height")
		photobg=rsvisual("background")
		photoBodyBack=rsvisual("BackBody")
		photoBodyFore=rsvisual("ForeBody")
		photoFg=rsvisual("Foreground")
		PhotoDate=rsvisual("adddate")
		isUpload=rsvisual("isupload")
		rsvisual.close
	case 2
		response.write "<form action=""z_visual_photo.asp"" method=post name=photo_confirm>"
		response.write "<input type=hidden name=""action"" value=""confirm"">"
		if request.cookies("myshow_"&userid)="" then
			response.write "<input type=hidden name=""uservisual"" value="""&v_myvisual&""">"
		else
			response.write "<input type=hidden name=""uservisual"" value="""&request.cookies("myshow_"&userid)&""">"
		end if
		photoid=checkstr(trim(request("photoid")))
		response.write "<input type=hidden name=""photoid"" value="""&photoid&""">"
		sqlvisual="select * from visual_photo where id="&photoid
		rsvisual.open sqlvisual,conn,1,1
		subtype="�յ�������"
		photoname=rsvisual("name")
		photomsg=rsvisual("content")
		fromuser=rsvisual("fromuser")
		photostatus=rsvisual("status")
		bgwidth=rsvisual("width")
		bgheight=rsvisual("height")
		photobg=rsvisual("background")
		photoBodyBack=rsvisual("BackBody")
		photoBodyFore=rsvisual("ForeBody")
		photoFg=rsvisual("Foreground")
		PhotoDate=rsvisual("adddate")
		isUpload=rsvisual("isupload")
		rsvisual.close
	case 3
		response.write "<form action=""z_visual_photo.asp"" method=post name=photo_design>"
		photoid=checkstr(trim(request("photoid")))
		response.write "<input type=hidden name=""photoid"" value="""&photoid&""">"
		sqlvisual="select * from visual_photo where id="&photoid
		rsvisual.open sqlvisual,conn,1,1
		subtype="ȷ�ϵ�����"
		photoname=rsvisual("name")
		photomsg=rsvisual("content")
		fromuser=rsvisual("fromuser")
		photostatus=rsvisual("status")
		bgwidth=rsvisual("width")
		bgheight=rsvisual("height")
		photobg=rsvisual("background")
		photoBodyBack=rsvisual("BackBody")
		photoBodyFore=rsvisual("ForeBody")
		photoFg=rsvisual("Foreground")
		FinishDate=rsvisual("FinishDate")
		PhotoDate=rsvisual("adddate")
		isUpload=rsvisual("isupload")
		if not isnull(FinishDate) then
			PhotoNameLeft=rsvisual("PhotoNameLeft")
			PhotoNameTop=rsvisual("PhotoNameTop")
			PhotoNameFont=rsvisual("PhotoNameFont")
			PhotoNameSize=rsvisual("PhotoNameSize")
			PhotoNameBold=rsvisual("PhotoNameBold")
			PhotoNameItalic=rsvisual("PhotoNameItalic")
			PhotoNameColor=rsvisual("PhotoNameColor")
			PhotoNameDirection=rsvisual("PhotoNameDirection")
			DateLeft=rsvisual("DateLeft")
			DateTop=rsvisual("DateTop")
			DateFont=rsvisual("DateFont")
			DateSize=rsvisual("DateSize")
			DateBold=rsvisual("DateBold")
			DateItalic=rsvisual("DateItalic")
			DateColor=rsvisual("DateColor")
			DateDirection=rsvisual("DateDirection")
		end if
		rsvisual.close
		if lcase(trim(membername))=lcase(trim(fromuser)) then		
			response.write "<input type=hidden name=""action"" value=""design"">"
			response.write "<input type=hidden name=""PhotoNameLeft"" value="""&PhotoNameLeft&""">"
			response.write "<input type=hidden name=""PhotoNameTop"" value="""&PhotoNameTop&""">"
			response.write "<input type=hidden name=""PhotoNameFont"" value="""&PhotoNameFont&""">"
			response.write "<input type=hidden name=""PhotoNameSize"" value="""&PhotoNameSize&""">"
			if PhotoNameBold then
				response.write "<input type=hidden name=""PhotoNameBold"" value=""1"">"
			else
				response.write "<input type=hidden name=""PhotoNameBold"" value=""1"">"
			end if
			if PhotoNameItalic then
				response.write "<input type=hidden name=""PhotoNameItalic"" value=""1"">"
			else
				response.write "<input type=hidden name=""PhotoNameItalic"" value=""1"">"
			end if
			response.write "<input type=hidden name=""PhotoNameColor"" value="""&PhotoNameColor&""">"
			response.write "<input type=hidden name=""PhotoNameDirection"" value="""&PhotoNameDirection&""">"
			response.write "<input type=hidden name=""DateLeft"" value="""&DateLeft&""">"
			response.write "<input type=hidden name=""DateTop"" value="""&DateTop&""">"
			response.write "<input type=hidden name=""DateFont"" value="""&DateFont&""">"
			response.write "<input type=hidden name=""DateSize"" value="""&DateSize&""">"
			if DateBold then
				response.write "<input type=hidden name=""DateBold"" value=""1"">"
			else
				response.write "<input type=hidden name=""DateBold"" value=""1"">"
			end if
			if DateItalic then
				response.write "<input type=hidden name=""DateItalic"" value=""1"">"
			else
				response.write "<input type=hidden name=""DateItalic"" value=""1"">"
			end if
			response.write "<input type=hidden name=""DateColor"" value="""&DateColor&""">"
			response.write "<input type=hidden name=""DateDirection"" value="""&DateDirection&""">"
		else
			response.write "<input type=hidden name=""action"" value=""accept"">"
		end if
	end select
	if isnull(FinishDate) then
		PhotoNameLeft=0
		PhotoNameTop=186
		PhotoNameFont="����"
		PhotoNameSize="9pt"
		PhotoNameBold=false
		PhotoNameItalic=false
		PhotoNameColor="#FFFFFF"
		PhotoNameDirection=0
		DateLeft=0
		DateTop=206
		DateFont="����"
		DateSize="9pt"
		DateBold=false
		DateItalic=false
		DateColor="#FFFFFF"
		DateDirection=0
	end if
	response.write "<table border=0 cellspacing=1 cellpadding=0 bgcolor="&Forum_Body(27)&" width=95% >"
	response.write "<tr height=25 valign=middle>"
	response.write "<th valign=middle colspan=2>"&forum_info(0)&"����� - "&subtype&"</th>"
	response.write "</tr>"
	response.write "<tr valign=middle height=30>"
	response.write "<td width=15% class=tablebody2 align=right>��Ӱ�����ƣ�</td>"
	response.write "<td align=left class=tablebody2>&nbsp;"
	if ListMethod=0 then
		response.write "<input type=text name=photoname size=80 maxlength=50>"
	else
		response.write "<b>"&photoname&"</b>"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr valign=middle height=90>"
	response.write "<td width=15% class=tablebody2 align=right>��������ݣ�</td>"
	response.write "<td align=left class=tablebody2>&nbsp;<textarea rows=5 cols=80 name=photomsg"
	if ListMethod<>0 then
		response.write " Readonly"
	end if
	response.write ">"&htmlencode(FilterJS(photomsg))&"</textarea></td>"
	response.write "</tr>"
	if ListMethod<>0 then
		response.write "<tr valign=middle height=30>"
		response.write "<td width=15% class=tablebody2 align=right>�����ߣ�</td>"
		response.write "<td align=left class=tablebody2>&nbsp;<a href=""dispuser.asp?name="&fromuser&""" target=_blank><b>"&fromuser&"</b></a></td>"
		response.write "</tr>"
	end if
	response.write "<tr valign=middle height=30>"
	response.write "<td width=15% class=tablebody2 align=right>�����ߣ�</td>"
	response.write "<td align=left class=tablebody2>&nbsp;"
	if ListMethod=0 then
		response.write "<input type=text name=photouser size=""50"">"
	else
		response.write "<b>"
		sqlvisual="select * from visual_photouser where photo_id="&photoid
		rsvisual.open sqlvisual,conn,1,1
		usercount=0
		do while not rsvisual.eof
			if rsvisual("confirmed") then
				if not isnull(rsvisual("uservisual")) and rsvisual("uservisual")<>"" then
					redim preserve uservisualsplit(usercount)
					uservisualsplit(usercount)=rsvisual("uservisual")
					redim preserve usernamesplit(usercount)
					usernamesplit(usercount)=rsvisual("username")
					if not isNull(FinishDate) then
						redim preserve LayerNo(UserCount)
						redim preserve OuterLeft(UserCount)
						redim preserve OuterTop(UserCount)
						redim preserve OuterWidth(UserCount)
						redim preserve OuterHeight(UserCount)
						redim preserve InnerLeft(UserCount)
						redim preserve InnerTop(UserCount)
						redim preserve InnerWidth(UserCount)
						redim preserve InnerHeight(UserCount)
						redim preserve Direction(UserCount)
						redim preserve NameLeft(UserCount)
						redim preserve NameTop(UserCount)
						redim preserve NameFont(UserCount)
						redim preserve NameSize(UserCount)
						redim preserve NameBold(UserCount)
						redim preserve NameItalic(UserCount)
						redim preserve NameColor(UserCount)
						redim preserve NameDirection(UserCount)
						LayerNo(UserCount)=rsvisual("LayerNo")
						OuterLeft(UserCount)=rsvisual("OuterLeft")
						OuterTop(UserCount)=rsvisual("OuterTop")
						OuterWidth(UserCount)=rsvisual("OuterWidth")
						OuterHeight(UserCount)=rsvisual("OuterHeight")
						InnerLeft(UserCount)=rsvisual("InnerLeft")
						InnerTop(UserCount)=rsvisual("InnerTop")
						InnerWidth(UserCount)=rsvisual("InnerWidth")
						InnerHeight(UserCount)=rsvisual("InnerHeight")
						Direction(UserCount)=rsvisual("Direction")
						NameLeft(UserCount)=rsvisual("NameLeft")
						NameTop(UserCount)=rsvisual("NameTop")
						NameFont(UserCount)=rsvisual("NameFont")
						NameSize(UserCount)=rsvisual("NameSize")
						NameBold(UserCount)=rsvisual("NameBold")
						NameItalic(UserCount)=rsvisual("NameItalic")
						NameColor(UserCount)=rsvisual("NameColor")
						NameDirection(UserCount)=rsvisual("NameDirection")
						if lcase(trim(rsvisual("UserName")))=lcase(trim(membername)) then hasFinished=rsvisual("Finished")
					end if
					usercount=usercount+1
				end if
			end if
			if lcase(trim(rsvisual("username")))<>lcase(trim(fromuser)) then
				response.write "<a href=""dispuser.asp?name="&rsvisual("username")&""" target=_blank>"&rsvisual("username")&"</a>"
				rsvisual.movenext
				if not rsvisual.eof then response.write ","
			else
				rsvisual.movenext
			end if
		loop
		rsvisual.close
		response.write "</b>"
	end if
	StepPos=70
	if bgwidth+70<140*usercount-(usercount-1)*StepPos then
		LeftPos=-35
		StepPos=(140*usercount-(bgwidth+70))/(usercount-1)
	else
		LeftPos=((bgwidth+70)-(140*usercount-(usercount-1)*StepPos))\2-35
	end if
	if isNull(FinishDate) then
		redim preserve LayerNo(UserCount-1)
		redim preserve OuterLeft(UserCount-1)
		redim preserve OuterTop(UserCount-1)
		redim preserve OuterWidth(UserCount-1)
		redim preserve OuterHeight(UserCount-1)
		redim preserve InnerLeft(UserCount-1)
		redim preserve InnerTop(UserCount-1)
		redim preserve InnerWidth(UserCount-1)
		redim preserve InnerHeight(UserCount-1)
		redim preserve Direction(UserCount-1)
		redim preserve NameLeft(UserCount-1)
		redim preserve NameTop(UserCount-1)
		redim preserve NameFont(UserCount-1)
		redim preserve NameSize(UserCount-1)
		redim preserve NameBold(UserCount-1)
		redim preserve NameItalic(UserCount-1)
		redim preserve NameColor(UserCount-1)
		redim preserve NameDirection(UserCount-1)
		for i=0 to usercount-1
			LayerNo(i)=i*20+10
			OuterLeft(i)=int(LeftPos+(140-StepPos)*i)
			OuterTop(i)=(bgheight-226)\2
			OuterWidth(i)=140
			OuterHeight(i)=226
			InnerLeft(i)=0
			InnerTop(i)=0
			InnerWidth(i)=140
			InnerHeight(i)=226
			Direction(i)=0
			NameLeft(i)=OuterLeft(i)+35
			NameTop(i)=0
			NameFont(i)="����"
			NameSize(i)="9pt"
			NameBold(i)=False
			NameItalic(i)=False
			NameColor(i)="#FFFFFF"
			NameDirection(i)=0
		next
	end if
	redim preserve UserVisualSplit(UserCount+5)
	redim preserve OuterLeft(UserCount+5)
	redim preserve OuterTop(UserCount+5)
	redim preserve OuterWidth(UserCount+5)
	redim preserve OuterHeight(UserCount+5)
	redim preserve InnerLeft(UserCount+5)
	redim preserve InnerTop(UserCount+5)
	redim preserve InnerWidth(UserCount+5)
	redim preserve InnerHeight(UserCount+5)
	redim preserve Direction(UserCount+5)
	for i=0 to 5
		UserVisualSplit(UserCount+i)=""
		OuterLeft(UserCount+i)=0
		OuterTop(UserCount+i)=0
		OuterWidth(UserCount+i)=140
		OuterHeight(UserCount+i)=226
		InnerLeft(UserCount+i)=0
		InnerTop(UserCount+i)=0
		InnerWidth(UserCount+i)=140
		InnerHeight(UserCount+i)=226
		Direction(UserCount+i)=0
	next
	if ListMethod=3 and not isNull(FinishDate) then
		dim rsPhotoUser
		set rsPhotoUser=server.createobject("ADODB.recordset")
		sqlvisual="select * from visual_Accouterment where photo_id="&PhotoID
		rsPhotoUser.open sqlvisual,conn,1,1
		do while not rsPhotoUser.eof
			i=rsPhotoUser("SeqNo")
			uservisualsplit(UserCount+i)=rsPhotoUser("ItemPicPath")
			OuterLeft(UserCount+i)=rsPhotoUser("OuterLeft")
			OuterTop(UserCount+i)=rsPhotoUser("OuterTop")
			OuterWidth(UserCount+i)=rsPhotoUser("OuterWidth")
			OuterHeight(UserCount+i)=rsPhotoUser("OuterHeight")
			InnerLeft(UserCount+i)=rsPhotoUser("InnerLeft")
			InnerTop(UserCount+i)=rsPhotoUser("InnerTop")
			InnerWidth(UserCount+i)=rsPhotoUser("InnerWidth")
			InnerHeight(UserCount+i)=rsPhotoUser("InnerHeight")
			Direction(UserCount+i)=rsPhotoUser("Direction")
			rsPhotoUser.movenext
		loop
		rsPhotoUser.close
		set rsPhotoUser=nothing
	end if
	if ListMethod=0 then
		response.write "&nbsp;"
		response.write "<SELECT name=friend onchange=DoTitle(this.options[this.selectedIndex].value)>"
		response.write "<OPTION selected value="""">ѡ��</OPTION>"
		sqlvisual="select F_friend from Friend where F_username='"&membername&"' and F_type=0 order by F_addtime desc"
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
			response.write "<OPTION value="""&rsvisual(0)&""">"&rsvisual(0)&"</OPTION>"
			rsvisual.movenext
		loop
		rsvisual.close
		response.write "</SELECT>"
	end if
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr valign=middle height=30>"
	response.write "<td width=15% class=tablebody2 align=right>��Ƭ��״̬��</td>"
	response.write "<td align=left class=tablebody2>&nbsp;"
	if ListMethod<>0 then
		response.write "<b>"
		if photostatus=1 then
			response.write "����"
		else
			response.write "����"
		end if
		response.write "</b>"
	else
		response.write "<input type=radio name=photostatus value=""0"">����&nbsp;&nbsp;<input type=radio name=photostatus value=""1"" checked>����"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name=photomyself value=""1"" onclick='if(this.checked){photo_request.photouser.disabled=true;photo_request.photouser.value="""";photo_request.friend.disabled=true;photo_request.photomsg.disabled=true;photo_request.photomsg.value="""";} else {photo_request.photouser.disabled=false;photo_request.friend.disabled=false;photo_request.photomsg.disabled=false;}'>����д��"
		if master then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name=forcephoto value=""1"">ǿ�ƺ�Ӱ"
		end if
	end if
	response.write "</td>"
	response.write "</tr>"
	if ListMethod=0 then
		response.write "<tr valign=middle height=30>"
		response.write "<td width=15% class=tablebody2 align=right>��Ч����</td>"
		response.write "<td align=left class=tablebody2><table width=100% cellpadding=0 cellspacing=0 border=0><tr><td width=50% class=tablebody2>"
		response.write "&nbsp;������"
		response.write "<select name=photobg onchange='selBg(this.options[this.selectedIndex].value)'>"
		response.write "<option value='' selected>�հױ���</option><option value='0'>�ҵı���</option><option value='1'>�ϴ��ı���</option>"
		if v_MyVip=1 then
			sqlvisual="select id,name from visual_items where content='1' and type=0 order by id"
		else
			sqlvisual="select id,name from visual_items where content='1' and type=0 and not vip order by id"
		end if
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
	  	if localpic=1 then
	  		response.write "<option value='1/"&right("0000"&rsvisual(0),4)&".GIF'>"&rsvisual(1)&"</option>"
	  	else
	  		response.write "<option value='"&right("00000000"&rsvisual(0),8)&"/01/00/'>"&rsvisual(1)&"</option>"
	  	end if
			rsvisual.movenext
		loop
		rsvisual.close
		response.write "</select></td>"
		response.write "<td width=50% class=tablebody2>&nbsp;���"
		response.write "<select name=photobodyback onchange='selBodyBack(this.options[this.selectedIndex].value)'>"
		response.write "<option value='' selected>�հ����</option>"
		if v_MyVip=1 then
			sqlvisual="select id,name from visual_items where content='2' and type=0 order by id"
		else
			sqlvisual="select id,name from visual_items where content='2' and type=0 and not vip order by id"
		end if
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
	  	if localpic=1 then
	  		response.write "<option value='2/"&right("0000"&rsvisual(0),4)&".GIF'>"&rsvisual(1)&"</option>"
	  	else
	  		response.write "<option value='"&right("00000000"&rsvisual(0),8)&"/02/00/'>"&rsvisual(1)&"</option>"
	  	end if
			rsvisual.movenext
		loop
		rsvisual.close
		response.write "</select></td></tr></table>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr valign=middle height=30>"
		response.write "<td width=15% class=tablebody2 align=right>ǰ��Ч����</td>"
		response.write "<td align=left class=tablebody2><table width=100% cellpadding=0 cellspacing=0 border=0><tr><td width=50% class=tablebody2>"
		response.write "&nbsp;ǰ����"
		response.write "<select name=photofg onchange='selFg(this.options[this.selectedIndex].value)'>"
		response.write "<option value='' selected>�հ�ǰ��</option>"
		if v_MyVip=1 then
			sqlvisual="select id,name from visual_items where content='25' and type=0 order by id"
		else
			sqlvisual="select id,name from visual_items where content='25' and type=0 and not vip order by id"
		end if
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
	  	if localpic=1 then
	  		response.write "<option value='25/"&right("0000"&rsvisual(0),4)&".GIF'>"&rsvisual(1)&"</option>"
	  	else
	  		response.write "<option value='"&right("00000000"&rsvisual(0),8)&"/25/00/'>"&rsvisual(1)&"</option>"
	  	end if
			rsvisual.movenext
		loop
		rsvisual.close
		response.write "</select></td>"
		response.write "<td width=50% class=tablebody2>&nbsp;��ǰ��"
		response.write "<select name=photobodyfore onchange='selBodyFore(this.options[this.selectedIndex].value)'>"
		response.write "<option value='' selected>�հ���ǰ</option>"
		if v_MyVip=1 then
			sqlvisual="select id,name from visual_items where content='24' and type=0 order by id"
		else
			sqlvisual="select id,name from visual_items where content='24' and type=0 and not vip order by id"
		end if
		rsvisual.open sqlvisual,conn,1,1
		do while not rsvisual.eof
	  	if localpic=1 then
	  		response.write "<option value='24/"&right("0000"&rsvisual(0),4)&".GIF'>"&rsvisual(1)&"</option>"
	  	else
	  		response.write "<option value='"&right("00000000"&rsvisual(0),8)&"/24/00/'>"&rsvisual(1)&"</option>"
	  	end if
			rsvisual.movenext
		loop
		rsvisual.close
		response.write "</select></td></tr></table>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr valign=middle height=30>"
		response.write "<td width=15% class=tablebody2 align=right>�ϴ�������</td>"
		response.write "<td align=left class=tablebody2>&nbsp;"
		response.write "<input type=hidden name=bgfname value="""">"
		response.write "<iframe frameborder=0 width=95% height=24 scrolling=no src=z_visual_upload.asp></iframe>"
		response.write "</td>"
	end if
	response.write "<tr valign=middle height=240>"
	response.write "<td colspan=2 class=tablebody2 align=center>"
	call ShowVisualPhoto(bgWidth,bgHeight,PhotoBg,PhotoBodyBack,PhotoBodyFore,PhotoFg,PhotoName,PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection,PhotoDate,DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection,UserCount,UserVisualSplit,LayerNo,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction,UserNameSplit,NameLeft,NameTop,NameFont,NameSize,NameBold,NameItalic,NameColor,NameDirection,True,isUpload,(UserCount=1 and ListMethod=3))
	response.write "</td>"
	response.write "</tr>"
	if ListMethod=0 then
		response.write "<tr valign=middle height=120>"
		response.write "<td class=tablebody2 colspan=2>&nbsp;<b>˵��</b>��<br>&nbsp;�� ������Ӣ��״̬�µĶ��Ž������߸���ʵ��Ⱥ�������<b>"&PhotoUserCount&"</b>���û����������û������Զ��޳�<br>&nbsp;�� �������<b>50</b>���ַ����������<b>"&GroupSetting(34)&"</b>���ַ�<br>&nbsp;�� ��Ӱ�������Ч��Ϊ<b>"&PhotoPeriod&"</b>�죬���ڵ����󽫱��Զ�ɾ��<br>&nbsp;�� �ϴ�������ͼ���ļ�����Ϊ280*226<br>&nbsp;<font color="&forum_body(8)&">�� ÿ����������ȡ����"
		if isnull(v_myvip) or v_myvip<>1 then
			response.write PhotoPrice1
		else
			response.write PhotoPrice2
		end if
		response.write "Ԫ(�������Լ�)�����������ֽ�"&mymoney&"Ԫ</font><br>&nbsp;�� �����������û�б�ͬ�⣬���з�����Ϊ����ϢȺ�����ÿ۳��������˻�</td>"
		response.write "</tr>"
	elseif ListMethod=3 then
		if lcase(trim(membername))=lcase(trim(fromuser)) then
			response.write "<tr valign=middle height=30>"
			response.write "<td colspan=2 class=tablebody2 align=center>"&vbNewLine
			response.write "�������<select name=SelBody onchange='selectBody(this.options[this.selectedIndex].value)'>"
			response.write "<option value=''>��ѡ������</option>"
			for i=0 to usercount-1
				response.write "<option value='"&i&"'>"&usernamesplit(i)&"</option>"
			next
			response.write "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "�������ƣ�<select name=SelName onchange='selectName(this.options[this.selectedIndex].value)'>"
			response.write "<option value=''>��ѡ������</option>"
			for i=0 to usercount-1
				response.write "<option value='"&i&"'>"&usernamesplit(i)&"</option>"
			next
			response.write "<option value='-1'>��Ӱ������</option>"
			response.write "<option value='-2'>��Ӱ��ʱ��</option>"
			response.write "</select>"
			if usercount>1 then
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������<select name=SelAccouterment onchange='selectAccouterment(this.options[this.selectedIndex].value)'>"
				response.write "<option value=''>��ѡ������</option>"
				response.write "<option value='0'>�������һ</option>"
				response.write "<option value='1'>��������</option>"
				response.write "<option value='2'>���������</option>"
				response.write "<option value='3'>��ǰ����һ</option>"
				response.write "<option value='4'>��ǰ�����</option>"
				response.write "<option value='5'>��ǰ������</option>"
				response.write "</select>"
			end if
			response.write "</td>"
			response.write "</tr>"
			response.write "<tr valign=middle height=30>"
			response.write "<td colspan=2 class=tablebody2 align=center>"
			response.write "������<input type=radio name=photo_mode value=0 checked disabled=true>�ƶ�&nbsp;<input type=radio name=photo_mode value=1 disabled=true>���ϼ���&nbsp;<input type=radio name=photo_mode value=2 disabled=true>���¼���&nbsp;<input type=radio name=photo_mode value=3 disabled=true>����&nbsp;<font color="&forum_body(8)&">ʹ��Ctrl�ͷ��������</font>"
			response.write "</td>"
			response.write "</tr>"
			response.write "<tr valign=middle height=30>"
			response.write "<td colspan=2 class=tablebody2 align=center>"
			if UserCount>1 then
				response.write "<select name=accouterment onchange='setAccouterment(this.options[this.selectedIndex].value)' disabled=true>"
				response.write "<option value=''>��ʹ������</option>"
				if v_myVip=1 then
					sqlvisual="select id,name,content from visual_items where type in ("&AccoutermentType&") order by id"
				else
					sqlvisual="select id,name,content from visual_items where type in ("&AccoutermentType&") and not vip order by id"
				end if
				rsvisual.open sqlvisual,conn,1,1
				do while not rsvisual.eof
			  	if localpic=1 then
			  		response.write "<option value='"&(split(rsvisual(2),"_")(0))&"/"&right("0000"&rsvisual(0),4)&".GIF'>"&rsvisual(1)&"</option>"
			  	else
			  		response.write "<option value='"&right("00000000"&rsvisual(0),8)&"/"&right("00"&(split(rsvisual(2),"_")(0)),2)&"/00/'>"&rsvisual(1)&"</option>"
			  	end if
					rsvisual.movenext
				loop
				rsvisual.close
				response.write "</select>&nbsp;&nbsp;"
			end if
			response.write "<input type=button name=photo_front value=""��ǰ"" disabled=true onclick='forward()'>&nbsp;<input type=button name=photo_rear value=""���"" disabled=true onclick='reverse()'>"
			response.write "&nbsp;&nbsp;��ת��<input type=radio name=rot_deg value=""0"" checked disabled=true onclick='selectRotation(this.value)'>0��&nbsp;<input type=radio name=rot_deg value=""1"" disabled=true onclick='selectRotation(this.value)'>90��&nbsp;<input type=radio name=rot_deg value=""2"" disabled=true onclick='selectRotation(this.value)'>180��&nbsp;<input type=radio name=rot_deg value=""3"" disabled=true onclick='selectRotation(this.value)'>270��&nbsp;&nbsp;<input type=checkbox name=mirror value=""0"" disabled=true onclick='selectMirror()'>��ת"
			response.write "</td>"
			response.write "</tr>"
			response.write "<tr valign=middle height=30>"
			response.write "<td colspan=2 class=tablebody2 align=center>"
			response.write "���壺<SELECT name=selFont onchange='selectFont(this.options[this.selectedIndex].value)' disabled=true>"
			response.write "<option value=""����"" selected>����</option>"
			response.write "<option value=""����_GB2312"">����</option>"
			response.write "<option value=""����_GB2312"">����</option>"
			response.write "<option value=""����"">����</option>"
			response.write "<option value=""����"">����</option>"
			response.write "<OPTION value=""Arial"">Arial</OPTION> "
			response.write "<OPTION value=""Arial Black"">Arial Black</OPTION>"
			response.write "<OPTION value=""Comic Sans MS"">Comic Sans MS</OPTION>"
			response.write "<OPTION value=""Courier New"">Courier New</OPTION>"
			response.write "<OPTION value=""Georgia"">Georgia</OPTION>"
			response.write "<OPTION value=""Impact"">Impact</OPTION>"
			response.write "<OPTION value=""Tahoma"">Tahoma</OPTION>"
			response.write "<OPTION value=""Times New Roman"">Times New Roman</OPTION>"
			response.write "<OPTION value=""Trebuchet MS"">Trebuchet MS</OPTION>"
			response.write "<OPTION value=""Script MT Bold"">Script MT Bold</OPTION>"
			response.write "<OPTION value=""Verdana"">Verdana</OPTION>"
			response.write "<OPTION value=""Lucida Console"">Lucida Console</OPTION>"
			response.write "</SELECT>&nbsp;&nbsp;"
			response.write "��С��<select name=selSize onChange='selectSize(this.options[this.selectedIndex].value)' disabled=true>"
			response.write "<option value=""6pt"">6pt</option>"
			response.write "<option value=""7pt"">7pt</option>"
			response.write "<option value=""8pt"">8pt</option>"
			response.write "<option value=""9pt"" selected>9pt</option>"
			response.write "<option value=""10pt"">10pt</option>"
			response.write "<option value=""11pt"">11pt</option>"
			response.write "<option value=""12pt"">12pt</option>"
			response.write "<option value=""14pt"">14pt</option>"
			response.write "<option value=""16pt"">16pt</option>"
			response.write "<option value=""18pt"">18pt</option>"
			response.write "<option value=""20pt"">20pt</option>"
			response.write "<option value=""24pt"">24pt</option>"
			response.write "</select>&nbsp;&nbsp;"
			response.write "��ɫ��<SELECT name=selColor onchange='selectColor(this.options[this.selectedIndex].value)' disabled=true>" 
			response.write "<option value="""">����</option>"
			response.write "<option style=""background-color:#F0F8FF;color: #F0F8FF"" value=""#F0F8FF"">#F0F8FF</option>"
			response.write "<option style=""background-color:#FAEBD7;color: #FAEBD7"" value=""#FAEBD7"">#FAEBD7</option>"
			response.write "<option style=""background-color:#00FFFF;color: #00FFFF"" value=""#00FFFF"">#00FFFF</option>"
			response.write "<option style=""background-color:#7FFFD4;color: #7FFFD4"" value=""#7FFFD4"">#7FFFD4</option>"
			response.write "<option style=""background-color:#F0FFFF;color: #F0FFFF"" value=""#F0FFFF"">#F0FFFF</option>"
			response.write "<option style=""background-color:#F5F5DC;color: #F5F5DC"" value=""#F5F5DC"">#F5F5DC</option>"
			response.write "<option style=""background-color:#FFE4C4;color: #FFE4C4"" value=""#FFE4C4"">#FFE4C4</option>"
			response.write "<option style=""background-color:#000000;color: #000000"" value=""#000000"">#000000</option>"
			response.write "<option style=""background-color:#FFEBCD;color: #FFEBCD"" value=""#FFEBCD"">#FFEBCD</option>"
			response.write "<option style=""background-color:#0000FF;color: #0000FF"" value=""#0000FF"">#0000FF</option>"
			response.write "<option style=""background-color:#8A2BE2;color: #8A2BE2"" value=""#8A2BE2"">#8A2BE2</option>"
			response.write "<option style=""background-color:#A52A2A;color: #A52A2A"" value=""#A52A2A"">#A52A2A</option>"
			response.write "<option style=""background-color:#DEB887;color: #DEB887"" value=""#DEB887"">#DEB887</option>"
			response.write "<option style=""background-color:#5F9EA0;color: #5F9EA0"" value=""#5F9EA0"">#5F9EA0</option>"
			response.write "<option style=""background-color:#7FFF00;color: #7FFF00"" value=""#7FFF00"">#7FFF00</option>"
			response.write "<option style=""background-color:#D2691E;color: #D2691E"" value=""#D2691E"">#D2691E</option>"
			response.write "<option style=""background-color:#FF7F50;color: #FF7F50"" value=""#FF7F50"">#FF7F50</option>"
			response.write "<option style=""background-color:#6495ED;color: #6495ED"" value=""#6495ED"">#6495ED</option>"
			response.write "<option style=""background-color:#FFF8DC;color: #FFF8DC"" value=""#FFF8DC"">#FFF8DC</option>"
			response.write "<option style=""background-color:#DC143C;color: #DC143C"" value=""#DC143C"">#DC143C</option>"
			response.write "<option style=""background-color:#00FFFF;color: #00FFFF"" value=""#00FFFF"">#00FFFF</option>"
			response.write "<option style=""background-color:#00008B;color: #00008B"" value=""#00008B"">#00008B</option>"
			response.write "<option style=""background-color:#008B8B;color: #008B8B"" value=""#008B8B"">#008B8B</option>"
			response.write "<option style=""background-color:#B8860B;color: #B8860B"" value=""#B8860B"">#B8860B</option>"
			response.write "<option style=""background-color:#A9A9A9;color: #A9A9A9"" value=""#A9A9A9"">#A9A9A9</option>"
			response.write "<option style=""background-color:#006400;color: #006400"" value=""#006400"">#006400</option>"
			response.write "<option style=""background-color:#BDB76B;color: #BDB76B"" value=""#BDB76B"">#BDB76B</option>"
			response.write "<option style=""background-color:#8B008B;color: #8B008B"" value=""#8B008B"">#8B008B</option>"
			response.write "<option style=""background-color:#556B2F;color: #556B2F"" value=""#556B2F"">#556B2F</option>"
			response.write "<option style=""background-color:#FF8C00;color: #FF8C00"" value=""#FF8C00"">#FF8C00</option>"
			response.write "<option style=""background-color:#9932CC;color: #9932CC"" value=""#9932CC"">#9932CC</option>"
			response.write "<option style=""background-color:#8B0000;color: #8B0000"" value=""#8B0000"">#8B0000</option>"
			response.write "<option style=""background-color:#E9967A;color: #E9967A"" value=""#E9967A"">#E9967A</option>"
			response.write "<option style=""background-color:#8FBC8F;color: #8FBC8F"" value=""#8FBC8F"">#8FBC8F</option>"
			response.write "<option style=""background-color:#483D8B;color: #483D8B"" value=""#483D8B"">#483D8B</option>"
			response.write "<option style=""background-color:#2F4F4F;color: #2F4F4F"" value=""#2F4F4F"">#2F4F4F</option>"
			response.write "<option style=""background-color:#00CED1;color: #00CED1"" value=""#00CED1"">#00CED1</option>"
			response.write "<option style=""background-color:#9400D3;color: #9400D3"" value=""#9400D3"">#9400D3</option>"
			response.write "<option style=""background-color:#FF1493;color: #FF1493"" value=""#FF1493"">#FF1493</option>"
			response.write "<option style=""background-color:#00BFFF;color: #00BFFF"" value=""#00BFFF"">#00BFFF</option>"
			response.write "<option style=""background-color:#696969;color: #696969"" value=""#696969"">#696969</option>"
			response.write "<option style=""background-color:#1E90FF;color: #1E90FF"" value=""#1E90FF"">#1E90FF</option>"
			response.write "<option style=""background-color:#B22222;color: #B22222"" value=""#B22222"">#B22222</option>"
			response.write "<option style=""background-color:#FFFAF0;color: #FFFAF0"" value=""#FFFAF0"">#FFFAF0</option>"
			response.write "<option style=""background-color:#228B22;color: #228B22"" value=""#228B22"">#228B22</option>"
			response.write "<option style=""background-color:#FF00FF;color: #FF00FF"" value=""#FF00FF"">#FF00FF</option>"
			response.write "<option style=""background-color:#DCDCDC;color: #DCDCDC"" value=""#DCDCDC"">#DCDCDC</option>"
			response.write "<option style=""background-color:#F8F8FF;color: #F8F8FF"" value=""#F8F8FF"">#F8F8FF</option>"
			response.write "<option style=""background-color:#FFD700;color: #FFD700"" value=""#FFD700"">#FFD700</option>"
			response.write "<option style=""background-color:#DAA520;color: #DAA520"" value=""#DAA520"">#DAA520</option>"
			response.write "<option style=""background-color:#808080;color: #808080"" value=""#808080"">#808080</option>"
			response.write "<option style=""background-color:#008000;color: #008000"" value=""#008000"">#008000</option>"
			response.write "<option style=""background-color:#ADFF2F;color: #ADFF2F"" value=""#ADFF2F"">#ADFF2F</option>"
			response.write "<option style=""background-color:#F0FFF0;color: #F0FFF0"" value=""#F0FFF0"">#F0FFF0</option>"
			response.write "<option style=""background-color:#FF69B4;color: #FF69B4"" value=""#FF69B4"">#FF69B4</option>"
			response.write "<option style=""background-color:#CD5C5C;color: #CD5C5C"" value=""#CD5C5C"">#CD5C5C</option>"
			response.write "<option style=""background-color:#4B0082;color: #4B0082"" value=""#4B0082"">#4B0082</option>"
			response.write "<option style=""background-color:#FFFFF0;color: #FFFFF0"" value=""#FFFFF0"">#FFFFF0</option>"
			response.write "<option style=""background-color:#F0E68C;color: #F0E68C"" value=""#F0E68C"">#F0E68C</option>"
			response.write "<option style=""background-color:#E6E6FA;color: #E6E6FA"" value=""#E6E6FA"">#E6E6FA</option>"
			response.write "<option style=""background-color:#FFF0F5;color: #FFF0F5"" value=""#FFF0F5"">#FFF0F5</option>"
			response.write "<option style=""background-color:#7CFC00;color: #7CFC00"" value=""#7CFC00"">#7CFC00</option>"
			response.write "<option style=""background-color:#FFFACD;color: #FFFACD"" value=""#FFFACD"">#FFFACD</option>"
			response.write "<option style=""background-color:#ADD8E6;color: #ADD8E6"" value=""#ADD8E6"">#ADD8E6</option>"
			response.write "<option style=""background-color:#F08080;color: #F08080"" value=""#F08080"">#F08080</option>"
			response.write "<option style=""background-color:#E0FFFF;color: #E0FFFF"" value=""#E0FFFF"">#E0FFFF</option>"
			response.write "<option style=""background-color:#FAFAD2;color: #FAFAD2"" value=""#FAFAD2"">#FAFAD2</option>"
			response.write "<option style=""background-color:#90EE90;color: #90EE90"" value=""#90EE90"">#90EE90</option>"
			response.write "<option style=""background-color:#D3D3D3;color: #D3D3D3"" value=""#D3D3D3"">#D3D3D3</option>"
			response.write "<option style=""background-color:#FFB6C1;color: #FFB6C1"" value=""#FFB6C1"">#FFB6C1</option>"
			response.write "<option style=""background-color:#FFA07A;color: #FFA07A"" value=""#FFA07A"">#FFA07A</option>"
			response.write "<option style=""background-color:#20B2AA;color: #20B2AA"" value=""#20B2AA"">#20B2AA</option>"
			response.write "<option style=""background-color:#87CEFA;color: #87CEFA"" value=""#87CEFA"">#87CEFA</option>"
			response.write "<option style=""background-color:#778899;color: #778899"" value=""#778899"">#778899</option>"
			response.write "<option style=""background-color:#B0C4DE;color: #B0C4DE"" value=""#B0C4DE"">#B0C4DE</option>"
			response.write "<option style=""background-color:#FFFFE0;color: #FFFFE0"" value=""#FFFFE0"">#FFFFE0</option>"
			response.write "<option style=""background-color:#00FF00;color: #00FF00"" value=""#00FF00"">#00FF00</option>"
			response.write "<option style=""background-color:#32CD32;color: #32CD32"" value=""#32CD32"">#32CD32</option>"
			response.write "<option style=""background-color:#FAF0E6;color: #FAF0E6"" value=""#FAF0E6"">#FAF0E6</option>"
			response.write "<option style=""background-color:#FF00FF;color: #FF00FF"" value=""#FF00FF"">#FF00FF</option>"
			response.write "<option style=""background-color:#800000;color: #800000"" value=""#800000"">#800000</option>"
			response.write "<option style=""background-color:#66CDAA;color: #66CDAA"" value=""#66CDAA"">#66CDAA</option>"
			response.write "<option style=""background-color:#0000CD;color: #0000CD"" value=""#0000CD"">#0000CD</option>"
			response.write "<option style=""background-color:#BA55D3;color: #BA55D3"" value=""#BA55D3"">#BA55D3</option>"
			response.write "<option style=""background-color:#9370DB;color: #9370DB"" value=""#9370DB"">#9370DB</option>"
			response.write "<option style=""background-color:#3CB371;color: #3CB371"" value=""#3CB371"">#3CB371</option>"
			response.write "<option style=""background-color:#7B68EE;color: #7B68EE"" value=""#7B68EE"">#7B68EE</option>"
			response.write "<option style=""background-color:#00FA9A;color: #00FA9A"" value=""#00FA9A"">#00FA9A</option>"
			response.write "<option style=""background-color:#48D1CC;color: #48D1CC"" value=""#48D1CC"">#48D1CC</option>"
			response.write "<option style=""background-color:#C71585;color: #C71585"" value=""#C71585"">#C71585</option>"
			response.write "<option style=""background-color:#191970;color: #191970"" value=""#191970"">#191970</option>"
			response.write "<option style=""background-color:#F5FFFA;color: #F5FFFA"" value=""#F5FFFA"">#F5FFFA</option>"
			response.write "<option style=""background-color:#FFE4E1;color: #FFE4E1"" value=""#FFE4E1"">#FFE4E1</option>"
			response.write "<option style=""background-color:#FFE4B5;color: #FFE4B5"" value=""#FFE4B5"">#FFE4B5</option>"
			response.write "<option style=""background-color:#FFDEAD;color: #FFDEAD"" value=""#FFDEAD"">#FFDEAD</option>"
			response.write "<option style=""background-color:#000080;color: #000080"" value=""#000080"">#000080</option>"
			response.write "<option style=""background-color:#FDF5E6;color: #FDF5E6"" value=""#FDF5E6"">#FDF5E6</option>"
			response.write "<option style=""background-color:#808000;color: #808000"" value=""#808000"">#808000</option>"
			response.write "<option style=""background-color:#6B8E23;color: #6B8E23"" value=""#6B8E23"">#6B8E23</option>"
			response.write "<option style=""background-color:#FFA500;color: #FFA500"" value=""#FFA500"">#FFA500</option>"
			response.write "<option style=""background-color:#FF4500;color: #FF4500"" value=""#FF4500"">#FF4500</option>"
			response.write "<option style=""background-color:#DA70D6;color: #DA70D6"" value=""#DA70D6"">#DA70D6</option>"
			response.write "<option style=""background-color:#EEE8AA;color: #EEE8AA"" value=""#EEE8AA"">#EEE8AA</option>"
			response.write "<option style=""background-color:#98FB98;color: #98FB98"" value=""#98FB98"">#98FB98</option>"
			response.write "<option style=""background-color:#AFEEEE;color: #AFEEEE"" value=""#AFEEEE"">#AFEEEE</option>"
			response.write "<option style=""background-color:#DB7093;color: #DB7093"" value=""#DB7093"">#DB7093</option>"
			response.write "<option style=""background-color:#FFEFD5;color: #FFEFD5"" value=""#FFEFD5"">#FFEFD5</option>"
			response.write "<option style=""background-color:#FFDAB9;color: #FFDAB9"" value=""#FFDAB9"">#FFDAB9</option>"
			response.write "<option style=""background-color:#CD853F;color: #CD853F"" value=""#CD853F"">#CD853F</option>"
			response.write "<option style=""background-color:#FFC0CB;color: #FFC0CB"" value=""#FFC0CB"">#FFC0CB</option>"
			response.write "<option style=""background-color:#DDA0DD;color: #DDA0DD"" value=""#DDA0DD"">#DDA0DD</option>"
			response.write "<option style=""background-color:#B0E0E6;color: #B0E0E6"" value=""#B0E0E6"">#B0E0E6</option>"
			response.write "<option style=""background-color:#800080;color: #800080"" value=""#800080"">#800080</option>"
			response.write "<option style=""background-color:#FF0000;color: #FF0000"" value=""#FF0000"">#FF0000</option>"
			response.write "<option style=""background-color:#BC8F8F;color: #BC8F8F"" value=""#BC8F8F"">#BC8F8F</option>"
			response.write "<option style=""background-color:#4169E1;color: #4169E1"" value=""#4169E1"">#4169E1</option>"
			response.write "<option style=""background-color:#8B4513;color: #8B4513"" value=""#8B4513"">#8B4513</option>"
			response.write "<option style=""background-color:#FA8072;color: #FA8072"" value=""#FA8072"">#FA8072</option>"
			response.write "<option style=""background-color:#F4A460;color: #F4A460"" value=""#F4A460"">#F4A460</option>"
			response.write "<option style=""background-color:#2E8B57;color: #2E8B57"" value=""#2E8B57"">#2E8B57</option>"
			response.write "<option style=""background-color:#FFF5EE;color: #FFF5EE"" value=""#FFF5EE"">#FFF5EE</option>"
			response.write "<option style=""background-color:#A0522D;color: #A0522D"" value=""#A0522D"">#A0522D</option>"
			response.write "<option style=""background-color:#C0C0C0;color: #C0C0C0"" value=""#C0C0C0"">#C0C0C0</option>"
			response.write "<option style=""background-color:#87CEEB;color: #87CEEB"" value=""#87CEEB"">#87CEEB</option>"
			response.write "<option style=""background-color:#6A5ACD;color: #6A5ACD"" value=""#6A5ACD"">#6A5ACD</option>"
			response.write "<option style=""background-color:#708090;color: #708090"" value=""#708090"">#708090</option>"
			response.write "<option style=""background-color:#FFFAFA;color: #FFFAFA"" value=""#FFFAFA"">#FFFAFA</option>"
			response.write "<option style=""background-color:#00FF7F;color: #00FF7F"" value=""#00FF7F"">#00FF7F</option>"
			response.write "<option style=""background-color:#4682B4;color: #4682B4"" value=""#4682B4"">#4682B4</option>"
			response.write "<option style=""background-color:#D2B48C;color: #D2B48C"" value=""#D2B48C"">#D2B48C</option>"
			response.write "<option style=""background-color:#008080;color: #008080"" value=""#008080"">#008080</option>"
			response.write "<option style=""background-color:#D8BFD8;color: #D8BFD8"" value=""#D8BFD8"">#D8BFD8</option>"
			response.write "<option style=""background-color:#FF6347;color: #FF6347"" value=""#FF6347"">#FF6347</option>"
			response.write "<option style=""background-color:#40E0D0;color: #40E0D0"" value=""#40E0D0"">#40E0D0</option>"
			response.write "<option style=""background-color:#EE82EE;color: #EE82EE"" value=""#EE82EE"">#EE82EE</option>"
			response.write "<option style=""background-color:#F5DEB3;color: #F5DEB3"" value=""#F5DEB3"">#F5DEB3</option>"
			response.write "<option style=""background-color:#FFFFFF;color: #FFFFFF"" value=""#FFFFFF"" selected>#FFFFFF</option>"
			response.write "<option style=""background-color:#F5F5F5;color: #F5F5F5"" value=""#F5F5F5"">#F5F5F5</option>"
			response.write "<option style=""background-color:#FFFF00;color: #FFFF00"" value=""#FFFF00"">#FFFF00</option>"
			response.write "<option style=""background-color:#9ACD32;color: #9ACD32"" value=""#9ACD32"">#9ACD32</option>"
			response.write "</SELECT>&nbsp;&nbsp;"
			response.write "<input type=checkbox name=selBold value=""1"" disabled=true onclick=""selectBold()"">����&nbsp;&nbsp;<input type=checkbox name=selItalic value=""1"" disabled=true onclick=""selectItalic()"">б��"
			response.write "</td>"
			response.write "</tr>"
		end if
	end if
	response.write "<tr valign=middle height=30>"
	response.write "<td class=tablebody2 valign=middle colspan=2 align=center>"
	select case ListMethod
	case 0
		response.write "<input type=Submit value=���� name=Submit>&nbsp; <input type=reset name=Clear value=��� onclick=""document.photo_bg.src='';document.photo_bg.style.display='none';document.photo_bodyback.src='';document.photo_bodyback.style.display='none';document.photo_bodyfore.src='';document.photo_bodyfore.style.display='none';document.photo_fg.src='';document.photo_fg.style.display='none';"">"
	case 1
		response.write "<a href=""z_visual.asp?shopid=297"">���ء������������б�</a>"
	case 2
		response.write "<input type=Submit value=ȷ�� name=Submit title=""ʹ�õ�ǰ������μ������Ӱ"">&nbsp; <input type=button name=goback value=���� onclick=""location.href='z_visual.asp?shopid=397'"" title=""���ء��յ��������б�"">"
	case 3
		if lcase(trim(membername))=lcase(trim(fromuser)) then
			response.write "<script language=javascript>userCount="""&UserCount&""";</script>"
			for i=0 to UserCount-1
				response.write "<input type=hidden name=""LayerNo_"&i&""" value="""&LayerNo(i)&""">"
				response.write "<input type=hidden name=""OuterLeft_"&i&""" value="""&OuterLeft(i)&""">"
				response.write "<input type=hidden name=""OuterTop_"&i&""" value="""&OuterTop(i)&""">"
				response.write "<input type=hidden name=""OuterWidth_"&i&""" value="""&OuterWidth(i)&""">"
				response.write "<input type=hidden name=""OuterHeight_"&i&""" value="""&OuterHeight(i)&""">"
				response.write "<input type=hidden name=""InnerLeft_"&i&""" value="""&InnerLeft(i)&""">"
				response.write "<input type=hidden name=""InnerTop_"&i&""" value="""&InnerTop(i)&""">"
				response.write "<input type=hidden name=""InnerWidth_"&i&""" value="""&InnerWidth(i)&""">"
				response.write "<input type=hidden name=""InnerHeight_"&i&""" value="""&InnerHeight(i)&""">"
				response.write "<input type=hidden name=""Direction_"&i&""" value="""&Direction(i)&""">"
				response.write "<input type=hidden name=""NameLeft_"&i&""" value="""&NameLeft(i)&""">"
				response.write "<input type=hidden name=""NameTop_"&i&""" value="""&NameTop(i)&""">"
				response.write "<input type=hidden name=""NameFont_"&i&""" value="""&NameFont(i)&""">"
				response.write "<input type=hidden name=""NameSize_"&i&""" value="""&NameSize(i)&""">"
				if NameBold(i) then
					response.write "<input type=hidden name=""NameBold_"&i&""" value=""1"">"
				else
					response.write "<input type=hidden name=""NameBold_"&i&""" value=""0"">"
				end if
				if NameItalic(i) then
					response.write "<input type=hidden name=""NameItalic_"&i&""" value=""1"">"
				else
					response.write "<input type=hidden name=""NameItalic_"&i&""" value=""0"">"
				end if
				response.write "<input type=hidden name=""NameColor_"&i&""" value="""&NameColor(i)&""">"
				response.write "<input type=hidden name=""NameDirection_"&i&""" value="""&NameDirection(i)&""">"
			next
			for i=0 to 5
				response.write "<input type=hidden name=""AccoutermentUserVisual_"&i&""" value="""&UserVisualSplit(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentOuterLeft_"&i&""" value="""&OuterLeft(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentOuterTop_"&i&""" value="""&OuterTop(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentOuterWidth_"&i&""" value="""&OuterWidth(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentOuterHeight_"&i&""" value="""&OuterHeight(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentInnerLeft_"&i&""" value="""&InnerLeft(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentInnerTop_"&i&""" value="""&InnerTop(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentInnerWidth_"&i&""" value="""&InnerWidth(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentInnerHeight_"&i&""" value="""&InnerHeight(UserCount+i)&""">"
				response.write "<input type=hidden name=""AccoutermentDirection_"&i&""" value="""&Direction(UserCount+i)&""">"
			next
			response.write "<input type=Submit value=ȷ�� name=Submit title=""ʹ�õ�ǰ�������Ϊ���յ���Ƭ"" onclick='SubmitDesign()'>&nbsp; "
		else
			if not isnull(FinishDate) and not hasFinished then
				response.write "<input type=Submit value=ͬ�� name=Submit title=""ͬ�⵱ǰ��Ƭ�����"">&nbsp; "
			end if
		end if
		response.write "<input type=button name=goback value=���� onclick=""location.href='z_visual.asp?shopid=497'"" title=""���ء�ȷ�ϵ������б�"">"
		if master and usercount>1 then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name=forcephoto value=""1"">ǿ��ͬ��"
		end if
	end select	
	response.write "</td>"
	response.write "</tr>"
	response.write "</table>"
	select case ListMethod
	case 0
		response.write "</form>"
	case 2
		response.write "</form>"
	case 3
		response.write "</form>"
	end select
end sub%>