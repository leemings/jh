<%
case 1
Application("yx8_mhjh_klt1")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>ͻȻһֻ��¹��͵Ϯ<font color=#0000FF>" & username & "</font>,��ȡ<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt1")*5000&"</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc1.asp?r="&Application("yx8_mhjh_klt1")&" target=optfrm><img src=img/tr003.gif  border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt1")*5000&" where ����='" & username & "'")

case 2
	mess="<font color=green>������ָ·��</font><font color=red>��������һλ�����޳��������������ˣ��Ϸ�����һ��Ѥ���Ĳʺ磬�ʺ��½����ã�������������챦��˭���õ�����Ȼ�����ǳ����ǰ;������û��"&jstl&"���������޷��ɹ��ģ�</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=mohuanshen.asp?r="&Application("yx8_mhjh_klt&r")&" target=optfrm><img src=img/mon18.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
case 3
Application("yx8_mhjh_klt3")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>ͻȻһֻ�Ӵ�Ŀֲ����޴���������ҧ��" & username & "����"&Application("yx8_mhjh_klt3")*2000&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc2.asp?r="&Application("yx8_mhjh_klt3")&" target=optfrm><img src=img/mon24.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt3")*2000&" where ����='" & username & "'")
case 4
	mess="<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center>�ʳǹ���</td> </tr> <tr><td valign=center align=center >ħ�ý����Ķ���ʱ����ÿСʱ�ĺ�30���ӣ�ǰ30���ӿ������ɴ���ͬ������������״̬����һ��ҪС�İ����𱻹���ҧ������ʹ�㿪���˱���Ҳ�޷��ֵ������Ϯ�������ԣ�һ��Ҫ���Լ���һЩ����������������û��Ҳһ��Ҫ���˵ģ������ݷֵ�״̬����,��Ҫ�������ϵ�����������Լ��ĸ���״̬.ħ��֮�񲻻�ӭ����Ŷ,�Ǻ�</td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table>"
case 5
Application("yx8_mhjh_klt4")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>������ͻȻһ���𵴣�һֻ�޴��ʬ�����Ƕ�������ȡ<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt4")*1000&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("yx8_mhjh_klt4")&" target=optfrm><img src=img/mon22.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt4")*1000&" where ����='" & username & "'")
case 6
Application("yx8_mhjh_klt4")=r
	mess="<font color=green>��������顿</font><font color=red>һ����͵ķ��������һֻ����ǧ��������Ȼ��������<font color=#0000FF>" & username & "</font>���ӻ���"&Application("yx8_mhjh_klt4")*20&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("yx8_mhjh_klt4")&" target=optfrm><img src=img/mon20.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt4")*20&" where ����='" & username & "'")
case 7
	mess="<font color=green>������͵Ϯ��</font><font color=red>һ�������ĺ����ɨ�������������Ⱥ�����Ű������<font color=#0000FF>" & username & "</font>��ҧȥ"&Application("yx8_mhjh_klt6")*100&"�������</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("yx8_mhjh_klt6")&" target=optfrm><img src=img/mon15.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt6")*100&" where ����='" & username & "'")
case 8
Application("yx8_mhjh_klt6")=r
	mess="<font color=green>������Ϯ����</font><font color=red>һ������ͻϮ��������������һȺ��Ů�ӵ��У����˾�ҧ��<font color=#0000FF>" & username & "</font>��ҧ������"&Application("yx8_mhjh_klt6")*2000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc5.asp?r="&Application("yx8_mhjh_klt6")&" target=optfrm><img src=img/mon23.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt6")*2000&" where ����='" & username & "'")
case 9
Application("yx8_mhjh_klt7")=r
	mess="<font color=green>������Ϯ����</font><font color=red>��¿�����ɽ��һ��Ұ��������ɽ���ε���ħ�ý�����������ʳ������<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt7")*500&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc7.asp?r="&Application("yx8_mhjh_klt7")&" target=optfrm><img src=img/mon28.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt7")*500&" where ����='" & username & "'")
case 10
Application("yx8_mhjh_klt8")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>ħ�ý����ڸ���ɽ������������һ�ɽ�������<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt8")*100&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc6.asp?r="&Application("yx8_mhjh_klt8")&" target=optfrm><img src=img/mon27.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt8")*100&" where ����='" & username & "'")
case 11
Application("yx8_mhjh_klt9")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>��~~~�ÿ��һֻѼ!�������Ӵ��ˣ�ʲô���У�һֻ����Ѽ�ܽ�ħ�ý�����͵��<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt9")*5000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc8.asp?r="&Application("yx8_mhjh_klt9")&" target=optfrm><img src=img/mon3.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt9")*5000&" where ����='" & username & "'")
case 12
Application("yx8_mhjh_klt10")=r
	mess="<font color=green>������͵Ϯ��</font><font color=red>������ͻȻ���ħ�ý�����һ��������<font color=#0000FF>" & username & "</font>��û���ں����ģ�����ȡ����"&Application("yx8_mhjh_klt10")*10&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc9.asp?r="&Application("yx8_mhjh_klt10")&" target=optfrm><img src=img/mon4.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt10")*10&" where ����='" & username & "'")
case 13
Application("yx8_mhjh_klt11")=r
	mess="<font color=green>��������顿</font><font color=red>һֻ���겻���Ľ���֮���ξ���ͻȻ�������������ڣ��͸�<font color=#0000FF>" & username & "</font>����"&Application("yx8_mhjh_klt11")*4000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc10.asp?r="&Application("yx8_mhjh_klt11")&" target=optfrm><img src=img/mon10.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+"&Application("yx8_mhjh_klt11")*4000&" where ����='" & username & "'")
case 14
Application("yx8_mhjh_klt12")=r
	mess="<font color=green>������Ϯ����</font><font color=red>һֻ��ʷ��ķ�ı���������ħ�ý�������<font color=#0000FF>" & username & "</font>�����Ƶ���ȡ����"&Application("yx8_mhjh_klt12")*1000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc11.asp?r="&Application("yx8_mhjh_klt12")&" target=optfrm><img src=img/mon9.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt12")*1000&" where ����='" & username & "'")
case 15
Application("yx8_mhjh_klt13")=r
	mess="<font color=green>����긽�塿</font><font color=red>����<font color=#0000FF>" & username & "</font>�����ڹ�ʱ��С���߻���ħ�������򱦱��Ļ��鸽�ŵ����ϣ���ҧ��"&Application("yx8_mhjh_klt13")*4000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc12.asp?r="&Application("yx8_mhjh_klt13")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt13")*4000&" where ����='" & username & "'")
case 16
Application("yx8_mhjh_klt14")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ���ﴳ��������!ɱ��" & username & "����"&Application("yx8_mhjh_klt14")*3000&"�㣡</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc4.asp?r="&Application("yx8_mhjh_klt14")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt14")*3000&" where ����='" & username & "'")
case 17
Application("yx8_mhjh_klt15")=r
	mess="<font color=green>������Ϯ����</font><font color=red>��Ư����~~~~һ����֪��ʲô�������������������,�˷��ж�ʧ��һ������,��" & username & "�õ���</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc15.asp?r="&Application("yx8_mhjh_klt15")&" target=optfrm><img src=img/mon12.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+1 where ����='" & username & "'")
case 18
Application("yx8_mhjh_klt16")=r
	mess="<font color=green>������Ϯ����</font><font color=red>��Ư������һ����֪��ʲô�������������������,�˷��ж�ʧ��һ������,��" & username & "�õ���</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc16.asp?r="&Application("yx8_mhjh_klt16")&" target=optfrm><img src=img/mon12.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����+1 where ����='" & username & "'")
case 19
Application("yx8_mhjh_klt17")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,��ȡ" & username & "����"&Application("yx8_mhjh_klt17")*500&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc17.asp?r="&Application("yx8_mhjh_klt17")&" target=optfrm><img src=img/Shenmo.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt17")*500&" where ����='" & username & "'")
case 20
Application("yx8_mhjh_klt18")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,���" & username & "����"&Application("yx8_mhjh_klt18")*10&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc18.asp?r="&Application("yx8_mhjh_klt18")&" target=optfrm><img src=img/Ying.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt18")*10&" where ����='" & username & "'")
case 21
Application("yx8_mhjh_klt19")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,���" & username & "����"&Application("yx8_mhjh_klt19")*10&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc19.asp?r="&Application("yx8_mhjh_klt&r")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt19")*10&" where ����='" & username & "'")
case 21
Application("yx8_mhjh_klt20")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,����" & username & "����"&Application("yx8_mhjh_klt20")*1200&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc20.asp?r="&Application("yx8_mhjh_klt20")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt20")*1200&" where ����='" & username & "'")
case 22
Application("yx8_mhjh_klt21")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,͵��" & username & "����"&Application("yx8_mhjh_klt21")*5&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc21.asp?r="&Application("yx8_mhjh_klt21")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt21")*5&" where ����='" & username & "'")
case 23
Application("yx8_mhjh_klt22")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,����" & username & "����"&Application("yx8_mhjh_klt22")*500&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc22.asp?r="&Application("yx8_mhjh_klt22")&" target=optfrm><img src=img/tr003.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt22")*500&" where ����='" & username & "'")
case 24
Application("yx8_mhjh_klt23")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,ɱ��" & username & "����"&Application("yx8_mhjh_klt23")*1900&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc23.asp?r="&Application("yx8_mhjh_klt23")&" target=optfrm><img src=img/tr002.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt23")*1900&" where ����='" & username & "'")
case 25
Application("yx8_mhjh_klt24")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,ɱ��" & username & "����"&Application("yx8_mhjh_klt24")*2300&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("yx8_mhjh_klt24")&" target=optfrm><img src=img/tr004.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt24")*2300&" where ����='" & username & "'")
case 26
Application("yx8_mhjh_klt25")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,ɱ��" & username & "����"&Application("yx8_mhjh_klt25")*2000&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc25.asp?r="&Application("yx8_mhjh_klt25")&" target=optfrm><img src=img/tr103.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt25")*2000&" where ����='" & username & "'")
case 27
Application("yx8_mhjh_klt26")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,ɱ��" & username & "����"&Application("yx8_mhjh_klt26")*1400&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc26.asp?r="&Application("yx8_mhjh_klt26")&" target=optfrm><img src=img/tr102.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt26")*1400&" where ����='" & username & "'")
case 28
Application("yx8_mhjh_klt27")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,ɱ��" & username & "����"&Application("yx8_mhjh_klt27")*1100&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc27.asp?r="&Application("yx8_mhjh_klt27")&" target=optfrm><img src=img/tr101.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt27")*1100&" where ����='" & username & "'")
case 29
Application("yx8_mhjh_klt28")=r
	mess="<font color=green>������Ϯ����</font><font color=red>ͻȻһֻ<font color=#0000FF>����</font>������,͵��" & username & "����"&Application("yx8_mhjh_klt28")*3000&"��</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc28.asp?r="&Application("yx8_mhjh_klt28")&" target=optfrm><img src=img/tr002.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update �û� set ����=����-"&Application("yx8_mhjh_klt28")*3000&" where ����='" & username & "'")
case 30
	mess="<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center>�ʳǹ���</td> </tr> <tr><td valign=center align=center >ħ�ý����Ķ���ʱ����ÿСʱ�ĺ�30���ӣ�ǰ30���ӿ������ɴ���ͬ������������״̬����һ��ҪС�İ����𱻹���ҧ������ʹ�㿪���˱���Ҳ�޷��ֵ������Ϯ�������ԣ�һ��Ҫ���Լ���һЩ����������������û��Ҳһ��Ҫ���˵�,�����ݷֵ�״̬����,��Ҫ�������ϵ�����������Լ��ĸ���״̬.ħ��֮�񲻻�ӭ����Ŷ,�Ǻǣ�</td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table>"
end select
conn.close
set conn=nothing
application.unlock
dati=Application("yx8_mhjh_talkarr")
		talkpoint=clng(Application("yx8_mhjh_talkpoint"))
		dim newdati(600)
		j=1
		for i=11 to 600 step 10
			newdati(j)=dati(i)
			newdati(j+1)=dati(i+1)
			newdati(j+2)=dati(i+2)
			newdati(j+3)=dati(i+3)
			newdati(j+4)=dati(i+4)
			newdati(j+5)=dati(i+5)
			newdati(j+6)=dati(i+6)
			newdati(j+7)=dati(i+7)
			newdati(j+8)=dati(i+8)
			newdati(j+9)=dati(i+9)
			j=j+10
		next
		newdati(591)=talkpoint+1
		newdati(592)=1
		newdati(593)=0
		newdati(594)=""
		newdati(595)="���"
		newdati(596)=""
		newdati(597)=namecolor
		newdati(598)=wordcolor
		newdati(599)=""&mess&""
		newdati(600)=chatroomsn
		Application.Lock
		Application("yx8_mhjh_talkpoint")=talkpoint+1
		Application("yx8_mhjh_talkarr")=newdati
		Application.UnLock
		erase newdati
		erase dati
end if
end if
%>