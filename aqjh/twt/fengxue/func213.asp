<%
function huayuan(name1,sh1,sy1)
titleline=1
%>
<%
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM �û� WHERE ����='" & session("aqjh_name") & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
huayuan=""
else
if application("huayuan")=0 then
news
randomize
rand=Int(62 * Rnd + 20)
Select Case rand
case "1"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�5000�����ӣ�"
sql="update �û� set ����=����-11,����=����+5000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "2"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�500�����ӣ�"
sql="update �û� set ����=����-11,����=����+500 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "3"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������һ��ǿ�У��ݶ�����ɱ����������ȴ��ȥ��3000���������⵽�ٸ�ͨ���������½�1000"
sql="update �û� set ����=����-11,����=����-3000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "4"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������׷�ϻ������ܣ�����1000�����ӣ�"
sql="update �û� set ����=����-11,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "5"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1000�����ӣ�"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "6"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "����׽�˼�ֻ��ȸ��������1000������2000��"
sql="update �û� set ����=����+2000,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "7"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��˼����100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "8"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��أ�����ȥ��֮�ʣ����صĻ��ش򿪣��㱻�Ҽ����ˣ���ʧ����10000��"
sql="update �û� set ����=����-10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "9"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��أ�����ȥ��֮�ʣ����صĻ��ش򿪣��㱻�Ҽ����ˣ���ʧ����10000��"
sql="update �û� set ����=����-10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "10"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�5000�����ӣ�"
sql="update �û� set ����=����-11,����=����+5000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "11"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�500�����ӣ�"
sql="update �û� set ����=����-11,����=����+500 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "12"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������һ��ǿ�У��ݶ�����ɱ����������ȴ��ȥ��3000���������⵽�ٸ�ͨ���������½�1000"
sql="update �û� set ����=����-11,����=����-3000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "13"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������׷�ϻ������ܣ�����1000�����ӣ�"
sql="update �û� set ����=����-11,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "14"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1000�����ӣ�"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "15"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "����׽�˼�ֻ��ȸ��������1000������2000��"
sql="update �û� set ����=����+2000,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "16"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��˼����100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "17"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��أ�����ȥ��֮�ʣ����صĻ��ش򿪣��㱻�Ҽ����ˣ���ʧ����10000��"
sql="update �û� set ����=����-10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "18"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "19"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "20"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��һ��С�ĵ�����ɽ�£��Ӵ˽�������һ������������"
sql="update �û� set ״̬='��',����=����-1000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�轣ʱ����ɽ�¶���','����')"
conn.execute sql
call boot(name1,"�轣ʱ����ɽ�¶���")
case "21"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "22"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ��ɫŮ�ӱ�ǿ��ǿ�ߣ���Ӣ�۾�������������1000��"
sql="update �û� set ����=����-100,����=����+10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "23"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "24"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ������ˣ���ָ�㣬���������2000��"
sql="update �û� set ����=����-11,����=����+2000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "25"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��ϴ��ϴ������������500"
sql="update �û� set ����=����-10,֪��=֪��+500 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "26"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱��˶��ڵ��ϵ�10�����ӣ�"
sql="update �û� set ����=����-11,����=����+10 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "27"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������������һ��ɫŮ�ӣ�������ȳ����ӣ���������5000��"
sql="update �û� set ����=����-1,����=����+5000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "28"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "29"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�10000�����ӣ�"
sql="update �û� set ����=����-11,����=����+10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "30"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������һ���󱦲ص�ͼ������һ��ð�գ���ɹ��õ���ʮ�򱦲ؼ��书��ѧ����������1000����"
sql="update �û� set ����=����+1000,����=����+500000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "31"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ������Ĳ˵�����ȥ���ˣ��õ�100����"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "32"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ�ֺóԵĸ�㣬��ȥ���õ�500����"
sql="update �û� set ����=����+11,����=����+500 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "33"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "34"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "35"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "36"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "37"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1000�����ӣ�"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "38"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "39"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "����͵�߱��˹�����·�,����300��."
sql="update �û� set ����=����-11,����=����+300 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close

case "40"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "41"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "42"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "43"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "����ɱ��һ����������������100"
sql="update �û� set ����=����-10,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "44"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˴󱦲أ������õ�;�б�һ��������͵Ϯ������"
sql="update �û� set ״̬='��',����=����-1000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�轣ʱ��͵Ϯ����','����')"
conn.execute sql
call boot(name1,"�轣ʱ��͵Ϯ����")
conn.close
case "45"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "46"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "47"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���������ߣ��ҿ��书��ǿ���龪һ����"
sql="update �û� set ����=����-11 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "48"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "49"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "50"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������и���Ҳ�ڴ˵أ��ڴ�30�غϺ��㲻�ж����ڳ�н��£�"
sql="update �û� set ״̬='��',����=����-1000,����=����-1000  where ����='" & session("aqjh_name") & "'"
conn.execute sql
sql="insert into into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','���轣�б�ɱ','����')"
conn.execute sql
call boot(name1,"���轣�б�ɱ")
conn.close
case "51"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "52"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1�������ӣ�"
sql="update �û� set ����=����-11,����=����+10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "53"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��ͻȻ��ǽ�Ƿ���1�����ӣ�"
sql="update �û� set ����=����-11,����=����+1 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "54"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���׷���һ��ʲôҲû��,������һ������.����ʧȥ1000"
sql="update �û� set ����=����-1000,����=����+0 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "55"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "56"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "57"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "58"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�300�����ӣ�"
sql="update �û� set ����=����-11,����=����+300 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "59"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�200�����ӣ�"
sql="update �û� set ����=����-11,����=����+200 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "60"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1000�����ӣ�"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "61"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "��������һ�����ߣ�֪������100"
sql="update �û� set ����=����-10,֪��=֪��+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "62"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "63"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "64"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "�������ذ��շ���׷��������ʧ5000��������͵����5000������"
sql="update �û� set ����=����-5000,����=����+5000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "65"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "66"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ������ˣ���ָ�㣬���������20000��"
sql="update �û� set ����=����-11,����=����+20000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "67"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "68"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "69"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "�����ϻ�ҧ�ˣ������½�11"
sql="update �û� set ����=����-11,����=����+0 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "70"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "71"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "�����������ٸ���Ա,˽�·�����10000���ŷ����뿪."
sql="update �û� set ����=����-11,����=����-10000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "72"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "73"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "����һ��ɫŮ��,ץȥ��������,׬��10000��,�����½�1000"
sql="update �û� set ����=����-11,����=����+10000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "74"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "������һ��ɫŮ�ӣ���ɫ�Ĵ��𣬶���ʩ���������½�1000��"
sql="update �û� set ����=����-1000,����=����-1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "75"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "76"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "77"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽���������µ�100������,��������,������һ����˽ӽ��ĽŲ���,��æ������"
sql="update �û� set ����=����-11,����=����+0 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "78"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�1000�����ӣ�"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "79"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "�������Ҳ�����������������뿪�ˡ�"
sql="update �û� set ״̬='����' where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "79"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "�������Ҳ�����������������뿪�ˡ�"
sql="update �û� set ״̬='����' where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "80"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢���˱���˽�ص�100�����ӣ�"
sql="update �û� set ����=����-11,����=����+100 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close

case "82"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢����һ��ʬ��,�����Ժ�õ�����1000"
sql="update �û� set ����=����-11,����=����+1000 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
case "1"
newer="" & session("aqjh_name") & "���轣��ʱ��" & sy1 & sh1 & "���㷢����һ���ر��⣬��ûԿ�ף��򲻿��ţ�"
sql="update �û� set ����=����-11 where ����='" & session("aqjh_name") & "'"
conn.execute sql
conn.close
End Select
end if
end if
huayuan=newer
end function
%>