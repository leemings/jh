<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=50){
window.alert("ˢǮ����ϲ�����𣿵㰡��ˢ������");
i=i+1;}top.location.href="../exit.asp"
}
</script>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fromname=LCase(trim(request.querystring("fromname")))
yn=LCase(trim(request.querystring("yn")))
toname=LCase(trim(request.querystring("toname")))
huaming=LCase(trim(request.querystring("huaming")))
xhsl=CInt(trim(request.querystring("wpsl")))
if InStr(fromname,"or")<>0 or InStr(fromname,"'")<>0 or InStr(fromname,"`")<>0 or InStr(fromname,"=")<>0 or InStr(fromname,"-")<>0 or InStr(fromname,",")<>0 then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End
end if
if toname<>aqjh_name then
	if fromname<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('������ʲô���˼�"&fromname&"Ҳ�����͸���ģ�');}</script>"
		Response.End
	else
		Response.Write "<script language=JavaScript>{alert('�в�ѽ"&fromname&"��������ʲô?�͸������Լ�ȥ�գ��������ģ�');}</script>"
		Response.End
	end if
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ա�,��ż,����,w7 from �û� where ����='"&fromname&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('�ͻ����˲����ڣ�������������Ϊ�����վ����ӳ��');}</script>"	'��4
	Response.End	
end if
if  mywpsl(rs("w7"),huaming)<xhsl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&fromname&"]����Ʒ��������("&xhsl&")����');}</script>"
	Response.End
end if
myhua=abate(rs("w7"),huaming,xhsl)
fromsex=rs("�Ա�")
peiou=rs("��ż")
qingren=rs("����")
conn.execute "update �û� set w7='"&myhua&"' where  ����='"&fromname&"'"
rs.close
rs.open "select �Ա� FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
tosex=rs("�Ա�")
randomize timer
r=int(rnd*3)
if yn=1 then
	if fromsex=tosex then
		if fromsex="��" then
			conn.execute "update �û� set ����=����-500 where ����='"&aqjh_name&"'"
			conn.execute "update �û� set ����=����-250 where ����='"&fromname&"'"
			Response.Write "<script language=JavaScript>{alert('������ʲô?����˸㲣����ϵѽ��̫�������ˣ�');}</script>"
			xianhua="���±�����["& aqjh_name &"]������{"& fromname &"}�͵�"& huaming &"�����������и㲣����ϵ�����˷��֣�"& aqjh_name &"�����½�500�㣬"& fromname &"�����½�250��!<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]���{"& fromname &"}�㲣����ϵ���˷��֣���������һƬǴ������������������Ϊ�����С���ˡ�����</marquee>"
		else
			conn.execute "update �û� set ����=����+200 where ����='"&aqjh_name&"'"
			conn.execute "update �û� set ����=����+100 where ����='"&fromname&"'"
			xianhua="��ϲ��["& aqjh_name &"]���˵ؽ�����{"& fromname &"}�͵�"& huaming &"��"& aqjh_name &"��������200�㣬"& fromname &"��������100��<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����{"& fromname &"}�͵��ʻ�����Ҿ���������Ů���Ӷ����и��ԣ����˵������������ˡ�����</marquee>"
		end if
	else
		conn.execute "update �û� set ����=����+"& xhsl*100 &" where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+"& xhsl*500 &" where ����='"&fromname&"'"
		xianhua="��ϲ��["& aqjh_name &"]���˵ؽ�����{"& fromname &"}�͵�"& huaming &"��"& aqjh_name &"��������"& 100*xhsl &"�㣬"& fromname &"��������"& 500*xhsl &"����Ĵ�Ϻ���ú����ۺ죬����������"
		if peiou=aqjh_name then
			select case r
				case 0
					if tosex="Ů" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�����Ϲ�{"& fromname &"}�͵��ʻ�������һ���Ҹ�ģ�����ÿɰ�ѽ��~~~~����ͷ�������ʻ����˿ɲ����ˣ���Ҿ���{"& fromname &"}�����Ǹ�ģ���ɷ�</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]��������{"& fromname &"}�͵��ʻ�������һ���Ҹ�ģ������������˯�����ˣ�~~~~����ͷ���Ϲ��ʻ����˿ɲ����ˣ���Ҿ���{"& fromname &"}�����Ǹ������ӣ�</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����{"& fromname &"}�͵��ʻ�������ס����ض�{"& fromname &"}˵���Ұ������ˣ��±������ǻ�Ҫ�����ޣ�</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�����Ϲ�{"& fromname &"}�͵��ʻ����������ˣ����ɵñ�ס{"& fromname &"}��ࣶ������ǵ�......������ͯ���ˣ�</marquee>"
			end select
		elseif qingren=aqjh_name then
			select case r
				case 0
					if tosex="Ů" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����������{"& fromname &"}�͵��ʻ�������һ���Ҹ�ģ�����ÿɰ�ѽ��~~~~��Ҿ���{"& fromname &"}�������е��Ǹ���������</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����С����{"& fromname &"}�͵��ʻ�������һ������ģ������������˯�����ˣ�~~~~����ͷ�������ӵ�Ҳ���ˣ��Ǻ�~~~~</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]��������{"& fromname &"}�͵��ʻ�������ס����ض�{"& fromname &"}˵���Ұ������ˣ��±�������Ҫ�����ޣ�</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]��������{"& fromname &"}�͵��ʻ����������ˣ����˲��ɵñ�ס��ࣶ������ǵ�......����Ҳ���¶Է����ɷ�/���ӿ��������Ǻù���Ү��</marquee>"
			end select
		else
			select case r
				case 0
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����{"& fromname &"}�͵��ʻ����˷ܹ��ȣ�����ε��ڵأ��٣�������������������ӷ�</marquee>"
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����{"& fromname &"}�͵��ʻ���{"& fromname &"}�˷ܵ�����ͨ�죬���۷��⣬���룺����������Ϸ�ˣ���</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]����{"& fromname &"}�͵��ʻ���["& aqjh_name &"]�����쳣�����ɵ���{"& fromname &"}���˼������ۣ����ǵ�......������Ҳ������</marquee>"
			end select
		end if
	end if
else
	if fromsex=tosex then
		if fromsex="��" then
			conn.execute "update �û� set ����=����+100 where ����='"&aqjh_name&"'"
			conn.execute "update �û� set ����=����-250,����=����-1000 where ����='"&fromname&"'"
			xianhua="�����±�������["& aqjh_name &"]�ܾ�����{"& fromname &"}�͵�"& huaming &"��"& aqjh_name &"��������100�㣬"& fromname &"�����½�250��!<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>{"& fromname &"}���["& aqjh_name &"]�㲣����ϵ���ܾ���["& aqjh_name &"]�����ض�{"& fromname &"}˵�����������������뷨�ǲ�������Ϊ���߿���������</marquee>"
		else
			conn.execute "update �û� set ����=����+100 where ����='"&aqjh_name&"'"
			conn.execute "update �û� set ����=����+50 where ����='"&fromname&"'"
			xianhua="�����±�������["& aqjh_name &"]�ܾ�����{"& fromname &"}�͵�"& huaming &"����Ҿ���������Ů���Ӷ�������˼��"& aqjh_name &"��������100�㣬"& fromname &"��������50��<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���{"& fromname &"}�͵��ʻ�����������˼��û����Ů���Ӹ�Ů�����ͻ�������</marquee>"
		end if
	else
		conn.execute "update �û� set ����=����+100 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����-1000 where ����='"&fromname&"'"
		xianhua="�����±�������["& aqjh_name &"]�ܾ�����{"& fromname &"}�͵�"& huaming &"��˵�����治����˼���Ҳ����������ʻ��������͸�������˰ɡ���<font color=red>"& aqjh_name &"��������100�㣬"& fromname &"��������1000 ��</font>"
		if peiou=aqjh_name then
			select case r
				case 0
					if tosex="Ů" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ����Ϲ�{"& fromname &"}�͵��ʻ�������ŭ�ݵ�˵��������������������������������㣬���ڲ������ֺ��ң�������û�У�����Ҿ���{"& fromname &"}�������е�̫�Ǹ�......��Ӧ�ã�</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ�������{"& fromname &"}�͵��ʻ��������㲵�˵���������ǻ������˺ã�������û������......������û��ɵ��۹⿴��{"& fromname &"}�����룺Ī����Ů��������......��</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ�����һ��{"& fromname &"}�͵��ʻ����е���ʹ�ض�{"& fromname &"}˵���������������ˣ����ǻ����ȷֿ�һ��ʱ�䣬���Է�һ�����ɵĿռ�ɣ��ÿ�������Ժ�Ҫ��ô��......��</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ�����һ��{"& fromname &"}�͵��ʻ����е�����ν�ض�{"& fromname &"}˵�����������Ѿ��޿�����ˣ���Ҳ�����˷����Ǯ�ˡ������Ϊ���������ģ���÷��޳�Ϊ���������Ӹе�ʹ��......</marquee>"
			end select
		elseif qingren=aqjh_name then
			select case r
				case 0
					if tosex="Ů" then
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���������{"& fromname &"}�͵��ʻ���Ц�ŵ��������������������ͷ���������ǿ����������ˡ�����Ҳ��ɵ������ֻ��ؿ���{"& fromname &"}����������......����������������㣬����Ҫ��ȥ����°���~~��</marquee>"
					else
						xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���С����{"& fromname &"}�͵��ʻ��������㲵�˵���������ǻ��Ƿ��ְ�......���ǿ��Ӷ���̫���ˣ���...������ÿ����ؿ���{"& fromname &"}�����룺��һ￴�������ڼ��ﵱ��ͷ�񣬹����������֣�Ҳ��������������</marquee>"
					end if
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ�������{"& fromname &"}�͵��ʻ�������ŭ�ݵ�˵�������������ˣ��ü��첻���ң������������ֺ��ң�������û�У�����Ҿ���{"& fromname &"}�����Ǻ�Ц����ã�������~~</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ�������{"& fromname &"}�͵��ʻ���Ц������˵�������ǲ���������һ������ô��ʱ��û�����ң������Ҳ���������......����Ҳ��ȥ����һ������</marquee>"
			end select
		else
			select case r
				case 0
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���{"& fromname &"}�͵��ʻ���������Ȼ��˵�������������ң�������Ѿ����������ˡ���</marquee>"
				case 1
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���{"& fromname &"}�͵��ʻ���Ц������˵�������������ѽ���ǲ�����ʲô��ı���ҿ������µ�Ӵ......��</marquee>"
				case else
					xianhua=xianhua&"<br><marquee width=100% scrollamount=8><font color=red><b>��С����Ϣ��</b></font>["& aqjh_name &"]�ܾ���{"& fromname &"}�͵��ʻ���һ����м�ؿ���{"& fromname &"}�������������ҵ���ͷ���ˣ������������ͷ��Ϲ���ۣ���</marquee>"
			end select
		end if
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>��������Ϣ��</b></font>"&xianhua	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>