<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"
n=request("n")
if n="" then n="��������"
%>
<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<style type=text/css>
A {text-decoration: none; color:'yellow'; font-size: 9pt;}
A:hover {text-decoration: none; color:ff6600;  font-size: 9pt;}
td {text-decoration: none; color:'yellow'; font-size: 9pt;}
</style>
</head>
<body leftmargin="0" topmargin="20" bgproperties="fixed" background=bg.gif>
<table width=100%><tbody><tr><td align=center>
<br><a href=?n=��������>����������</a><br>
<a href=?n=��������>�����������</a><br>
<a href=?n=��������>�󽭺����С�</a><br>
<a href=?n=��������>���������ˡ�</a><br>
<a href=?n=�ٸ�����>��ٸ����š�</a><br>
<a href=?n=��Աר��>���Աר����</a><br>
<a href=?n=��������>���������ܡ�</a><br><br>
<%
select case n
case "��������"
%>
<table width=120>
<tr><td><a href=# onClick="window.open('../hcjs/jhzb/index.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">�ҵ�װ��</A></td><td><a href=# onclick="window.open('../hcjs/jhzb/myhua.asp','aqjh_win','scrollbars=0,resizable=0,width=430,height=340')">�ҵ��ʻ�</a></td></tr>
<tr><td><a href=# onClick="window.open('../JHMP/MP5.ASP','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">�ҵ�ͽ��</a></td><td><a href=# onClick="window.open('../hcjs/box/index.asp','aqjh_win','scrollbars=1,resizable=1,width=760,height=400')">�ҵĴ���</a></td</tr>
<tr><td><a href=# onClick="window.open('../yamen/regmodi.asp','aqjh_win','scrollbars=1,resizable=0,width=480,height=470')">�����޸�</a></td><td><a href=# onClick="window.open('../yamen/editpass2.asp','aqjh_win','scrollbars=1,resizable=0,width=520,height=340')">���뱣��</A></td></tr>
<tr><td><a href=# onClick="window.open('../yamen/modify.asp','aqjh_win','scrollbars=1,resizable=0,width=520,height=340')">�޸�����</a></td><td><a href=# onClick="window.open('../hcjs/jhmt/','aqjh_win','width=550,height=410,resizable=0,scrollbars=0')">������̽</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/wuguan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">��������</a></td><td><a href=# onClick="window.open('../HCJS/zstt/zs.asp','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">ת��Ͷ̥</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/money.asp','aqjh_win','scrollbars=0,resizable=0,width=430,height=440')">��ȡ����</a></td><td><a href=# onclick="window.open('../work/changeworkzhongzu.asp','aqjh_win','scrollbars=1,resizable=1,width=400,height=300')">��������</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhmp/myfriend.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')">���˼�¼</a></td><td><a href=# onClick="window.open('../hcjs/sendphoto/PHOTO.ASP','aqjh_win','scrollbars=1,resizable=1,width=750,height=500')">������Ƭ</a></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/long/long.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')">�����̵�</a></td><td><a href=# onClick="window.open('../hcjs/long/longeat.asp','aqjh_win','scrollbars=1,resizable=1,width=750,height=500')">��������</a></td></tr>

<tr>
	<td><a href='setfontsize.asp'>��������</a></td>
	<td><a href='face.asp'>ͷ������</a></td>
</tr>
<tr>
	<td><a href=# onClick="window.open('webicq.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')">�ɸ봫��</a></td></td>
	<td><a href=../help.asp target=_blank>����</a></td>
</tr>
</table>
<%
case "��������"
%>
<table width=120>
<tr><td><a href=# onClick="window.open('../WORK/CATCH/CATCH.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">������˽</A></td><td><a href=# onClick="window.open('../twt/hit/DIAOYU.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��������</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/hunyin/killer.asp','aqjh_win','scrollbars=1,resizable=0,width=640,height=480')">��Ӷɱ��</A></td><td><a href=# onClick="window.open('../twt/fengxue/fengxue.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">��ѩ�轣</a></td></tr>
<tr><td><a href=# onClick="window.open('../twt/findbao/','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">�µ�ð��</A></td><td><a href=# onClick="window.open('../twt/diaoyu/diaoyu.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">���е���</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/grade/DIAOYU.ASP','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">����Ѱ��</A></td><td><a href=# onClick="window.open('../twt/wokou/index.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">��������</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/loot/loot.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">��������</A></td><td><a href=# onClick="window.open('../twt/hunyin/stunt.asp','aqjh_win','scrollbars=1,resizable=1,width=580,height=350')">�����ؼ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../work/tie/tiemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=370')">��ʯ����</a></td><td><a href=# onClick="window.open('../work/mine/minemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=470')">�����ɿ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../work/famu/famumain.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=370')">������ľ</A></td><td><a href=# onClick="window.open('../work/ice/icemain.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=430')">�����ɱ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/wabao/wabao.asp','aqjh_win','scrollbars=1,resizable=0,width=450,height=340')">�ھ򱦲�</A></td><td><a href=# onclick="window.open('../hcjs/dzwq/index.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')" >��������</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/peiyao/peiyao.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')">��ҩϵͳ</A></td><td><a href=# onClick="window.open('../twt/tao/index.asp','aqjh_win','scrollbars=0,resizable=0,width=780,height=460')">�����ͽ�</a></td></tr>
<tr><td><A href="../TOP/TOP47.ASP" target=_blank>��������</A></td><td><a href="#"  onClick="javascript:chatwin=window.open('../top/top.htm','','scrollbars=no,resizable=no,width=760,height=400,top=0,left=0')">��������</a></td></tr>
<tr><td><a href=# onClick="window.open('../HCJS/bang/meeting.asp','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')">���ִ��</a></td><td><a href=../jhshow/jl.asp target=_blank><font color="#FF0000">��װ����</a></td></tr>
<tr><td><a href=../jhshow/ target=_blank><font color="#FF0000">��������</a></td><td><a href=# onClick="window.open('../garden/','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')"><font color="#FF0000">��԰����</a></td></tr>
</table>
<%
case "��������"
%>
<table width=120>
<tr><td><a href="../CHANGAN/xiaowu.asp" target=_blank>����С��</A></td><td><a href=# onClick="window.open('../hcjs/qr/index.asp','aqjh_win','scrollbars=1,resizable=0,width=650,height=450')">����С��</A></td></tr><tr><td><a href=# onClick="window.open('../jhjd/jd.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')">�����Ƶ�</A></td><td><a href=# onClick="window.open('../hcjs/pub/pub.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">���ſ�ջ</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/jhjs/hua.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=460')"">��������</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/wuqi.asp','aqjh_win','scrollbars=0,resizable=0,width=650,height=550')">�����̵�</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/jhjs/anqi.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=400')">�����̵�</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/yaopu.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">����ҩ��</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/jhjs/vehicle.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">��ͨ����</A></td><td><a href=# onClick="window.open('../hcjs/jhjs/cw.asp','aqjh_win','scrollbars=0,resizable=0,width=520,height=450')">�����̵�</A></td></tr><tr><td><a href=# onClick="window.open('../jhmp/maidou.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')">��������</a></td><td><a href=# onClick="window.open('../hcjs/yilao.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">����ɽׯ</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/yhy/','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">������Ժ</A></td><td><a href=# onClick="window.open('../hcjs/yhyb/index.asp','aqjh_win','scrollbars=1,resizable=1,width=780,height=460')">����Ѽ��</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/qr/fkyy.asp','aqjh_win','scrollbars=1,resizable=0')">��������</A></td><td><a href=# onClick="window.open('../yamen/yiyuan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">����ҽԺ</A></td></tr><tr><td><a href=# onClick="window.open('../garden/hua.htm','aqjh_win','scrollbars=0,resizable=0,width=670,height=400')">������԰</a></td><td><a href=# onClick="window.open('../hcjs/jhjs/dan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��������</A></td></tr>
<tr><td><a href=# onClick="window.open('../CHANGAN/fangwu.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=450')">��������</A></td><td><a href=# onClick="window.open('../hcjs/mj/mujuan.asp','aqjh_win','scrollbars=1,resizable=0,width=690,height=440')">����ļ��</A></td></tr><tr><td><a href=# onClick="window.open('../jhmp/gry.asp','aqjh_win','scrollbars=0,resizable=0,width=350,height=230')">�����¶�</A></td><td><a href=# onClick="window.open('../gupiao/stock.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��Ʊ�г�</A></td></tr><tr><td><a href=# onClick="window.open('../hcjs/dk/Daikuan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��������</A></td><td><a href=# onClick="window.open('../hcjs/wish/wish.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">������</A></td></tr>
<tr><td><a href=# onClick="window.open('../twt/qq/read.htm','aqjh_win','scrollbars=1,resizable=0,width=540,height=480')">��������</A></td><td><a href=# onClick="window.open('../work/changework.asp','aqjh_win','scrollbars=0,resizable=0,width=520,height=380')">ְҵת��</A></td></tr><tr><td><a href=# onClick="window.open('../twt/hunyin/','aqjh_win','scrollbars=0,resizable=0,width=550,height=450')">��������</A></td><td><a href=# onClick="window.open('../twt/hunyin/taohun.asp','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">�����ӻ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../bet/betindex.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">�����Ĺ�</A></td><td><a href=# onClick="window.open('../hcjs/mc/','aqjh_win','scrollbars=1,resizable=1,width=670,height=400')">������</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/xuewu/xuetang.htm','aqjh_win','scrollbars=1,resizable=1,width=640,height=380')">���ϰ��</A></td><td><a href=# onClick="window.open('../twt/dg/dg.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=400')">������</a></td></tr><tr><td><a href=# onClick="window.open('../hcjs/zysc/zysc.asp','aqjh_win','scrollbars=1,resizable=1,width=700,height=400')"><font color="#FF0000">�����г�</a></td><td><a href=# onClick="window.open('../xcj/zh/','aqjh_win','scrollbars=1,resizable=1,width=500,height=300')"><font color="#FF0000">ת���г�</a></td></tr>
</table>
<%
case "��������"
%>
<table width=120>
<tr align=center><td><a href=# onClick="window.open('../jhmp/index.asp','aqjh_win','scrollbars=1,resizable=0,width=760,height=450')">���Ű���</A></td><td><a href=# onClick="window.open('../jhmp/setwg.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">�书����</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhmp/addmp/chuangp.asp','aqjh_win','scrollbars=1,resizable=0,width=550,height=450')">��������</a></td><td><a href=# onClick="window.open('../jhmp/addmp/rangwei.asp','aqjh_win','scrollbars=1,resizable=1,width=580,height=350')">������λ</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../hcjs/room/','aqjh_win','scrollbars=1,resizable=0,width=600,height=450')">�Դ�����</a></td><td><a href=# onclick="window.open('../jhmp/mympjj.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')" target=_self>���ɻ���</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/mpxs/mpxs.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=450')">����нˮ</a></td><td><a href=# onclick="window.open('../xcj/mmxl/index.asp','aqjh_win','scrollbars=1,resizable=1,width=320,height=340')" target=_self>��������</A></td></tr>
</table>
<%
case "�ٸ�����"
%>
<table width=120>
<tr align=center><td><a href=# onClick="window.open('../yamen/admin.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')">��������</A></td><td><a href=# onClick="window.open('../yamen/laofan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/minan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=470')">������ѯ</A></td><td><a href=# onClick="window.open('../jl/jlmain.asp','aqjh_win','scrollbars=1,resizable=0,width=740,height=450')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/laofan1.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">˯�߽��</A></td><td><a href=# onClick="window.open('../hcjs/sy/sy.htm','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">������ԩ</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../yamen/laofang2.asp','aqjh_win','scrollbars=0,resizable=0,width=620,height=400')">�鿴���</A></td><td><a href=../bbs/ target=_blank><font color=red>������̳</a></td></tr>
</table>
<%
case "��Աר��"
%>
<table width=120>
<tr><td><a href="hy.asp" target=_blank>�����Ա</A></td><td><a href=# onClick="window.open('../jhmp/hy.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">��Ա�б�</A></td></tr>
<tr><td><a href=# onClick="window.open('../hcjs/card/card.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��Ա��Ƭ</A></td><td><a href=# onClick="window.open('../jhhy/workauto/ice/icemain.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=430')">��Ա�ɱ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/hyxuewu/','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')">��Աѧ��</a></td><td><a href=# onClick="window.open('../jhhy/xiang/index.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">��Ա����</a></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/workauto/famu/famumain.asp','aqjh_win','scrollbars=1,resizable=0,width=600,height=370')">��Ա��ľ</A></td><td><a href=# onClick="window.open('../jhhy/workauto/mine/minemain.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=450')">��Ա�ɿ�</A></td></tr>
<tr><td><a href=# onClick="window.open('../jhhy/hywabao/wabao.asp','aqjh_win','scrollbars=1,resizable=0,width=500,height=450')">��Ա�ڱ�</a></td><td><a href=# onClick="window.open('../jhhy/hyjb/index.asp','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')">��Ա����</a></td></tr>
</table>
<%
case "��������"
%>
<table width=120>
<tr align=center><td><a href="face.asp" scrollbars="0" resizable="0" width="740" height="450">����ͷ��</A></td><td><a href=# onClick="window.open('../liwu/xrlw.asp','f3')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/yxs/yxs.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=470')">��н��ȡ</A></td><td><a href=# onClick="window.open('../liwu/zllw.asp','f3')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhmp/bxmp.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">��������</A></td><td><a href=# onClick="window.open('../liwu/zwlw.asp','f3')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../xcj/xgn/sm.ASP','aqjh_win','scrollbars=0,resizable=0,width=620,height=400')">����˵��</A></td><td><a href=../bbs/ target=_blank><font color=red>������̳</a></td></tr>
<tr align=center><td><a href=# onClick="window.open('../hcjs/pig/zhu.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">�����¸�</A></td><td><a href=# onClick="window.open('jhjh/index.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">����ϵͳ</A>
<tr align=center><td><a href=# onClick="window.open('../liwu/zmlw.asp','f3')">��ĩ��Ʒ</A></td><td><a href=# onClick="window.open('../liwu/zzlw.asp','f3')">վ������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../liwu/ymlw.asp','f3')">��ĩ����</A></td><td><a href=# onClick="window.open('../liwu/jrlw.asp','f3')">��������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('/twt/mhmj/mhmj/index.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')">�λ�ħ��</A></td><td><a href=# onClick="window.open('../hcjs/zstt/yld.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">�������</A></td></tr>
<tr align=center><td><a href=# onClick="window.open('../jhhy/hyzh/zhuan.asp','aqjh_win','scrollbars=1,resizable=0,width=620,height=400')">ת����Ա</A></td><td><a href=# onClick="window.open('../jhhy/ydxx/DAJIA.ASP','aqjh_win','scrollbars=0,resizable=0,width=700,height=420')">�컯����</A></td></tr>

</A></td></tr><tr align=center><td></td>
</table>
<%
end select
%>
</td></tr></tbody></table> 