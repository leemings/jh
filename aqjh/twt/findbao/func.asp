<%
function huayuan(name1,sh1,sy1)
%> <!--#include file="data.asp"--> <%
sql="SELECT * FROM �û� WHERE ����='" & name1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
      huayuan=""
else
      if rs("����")=Application("aqjh_baowuname") then
            huayuan="" & name1 & "���Ѿ���"& Application("aqjh_baowuname") &"�˱��������Ժ���ܵ��µ�Ѱ����" 
      else
            randomize timer
            rand=Int(79 * Rnd + 10) 
                 if rand>=10 and rand<=50 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�100�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+100,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=51 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�5000�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+5000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=52 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�500�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+500,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=53 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "����������Ұ�ޣ���ս֮���������Ұ�ޣ���ȴ����3000������"
                        sql="update �û� set ����=����-11,����=����-3000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=54 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��������ʯͷһ��ˤ��һ�ӣ�����100�����ӣ�"
                        sql="update �û� set ����=����-11,����=����-100,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=55 or rand=85 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢����һ���ص�ɽ��ҩ��һ��֮�£����о綾������"                         
                        sql="update �û� set ״̬='��',����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�ڹµ��гԴ��ж�ʳ�����','����')"
                        'sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�ڹµ��гԴ��ж�ʳ�����')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=56 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˹Ŵ����أ�����ȥ��֮�ʣ����صĻ��ش򿪣��㱻�Ҽ����ˣ���ʧ����10000��"
                        sql="update �û� set ����=����-10000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=57 or rand=86 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "����ʧ��������£��Ӵ˽�������һ������������"
                        sql="update �û� set ״̬='��' where ����='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�ڹµ�ʧ���������','����')"
                        'sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�ڹµ�ʧ���������')"
                        conn.execute sql
                      call boot(name1)
                 elseif rand=58 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������һ��ɫŮ�ӱ��ϻ�׷�ϣ���Ӣ�۾�������������1000��"
                        sql="update �û� set ����=����-100,����=����+1000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=59 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������һ������ˣ���ָ�㣬���������200��"
                        sql="update �û� set ����=����-11,����=����+200,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=60 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��ϴ��ϴ������������50"
                        sql="update �û� set ����=����-10,����=����+50 where,����ʱ��=now() ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=61 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������һ��ɫŮ�ӣ�������ȳ��µ�����������500��"
                        sql="update �û� set ����=����-1,����=����+500,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=62 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��������һ���󱦲أ�����һ��ð�գ���ɹ��õ���ʮ�򱦲ؼ��书��ѧ,��������1000��"
                        sql="update �û� set ����=����+1000,����=����+500000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand>=63 or rand<=66 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�1000�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+1000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=67 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��ϴ��ϴ������������100"
                        sql="update �û� set ����=����-10,����=����+100,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=68 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˴󱦲أ������õ�;�б��Ҽ�������"
                        sql="update �û� set ״̬='��' where ����='" & name1 & "'"
                        conn.execute sql 
                       sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�µ����ñ���ʱ���Ҽ�����','����')" 
                      '  sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�µ����ñ���ʱ���Ҽ�����')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=69 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���������ߣ��ҿ��书��ǿ���龪һ����"
                        sql="update �û� set ����=����-11 where,����ʱ��=now() ����='" & name1 & "'"
conn.execute sql
                 elseif rand=70 or rand=87 then
huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��������и���Ҳ�ڴ˵أ��ڴ�30�غϺ��㲻�ж����ڳ�н��£�"
                        sql="update �û� set ״̬='��' where ����='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�ڹµ�ð���б�ɱ','����')" 
                       ' sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�ڹµ�ð���б�ɱ')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=71 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�1�������ӣ�"
                        sql="update �û� set ����=����-11,����=����+10000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=72 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��ͻȻ��ǽ�Ƿ���1�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+1,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=73 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢����һ�����ص�ɽ��ҩ��һ��֮�£���������2000��"
                        sql="update �û� set ����=����-11,����=����+2000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=74 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�300�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+300,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=75 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢���˱������µ�200�����ӣ�"
                        sql="update �û� set ����=����-11,����=����+200,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=76 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��ϴ��ϴ������������100"
                        sql="update �û� set ����=����-10,����=����+100,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=77 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "��ʧ�����ɽ�£����Ѳ�����������ʧ5000�������Ѳ��������к󸣣�����5000������"
                        sql="update �û� set ����=����-5000,����=����+5000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=78 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������һ������ˣ���ָ�㣬���������2000��"
                        sql="update �û� set ����=����-11,����=����+2000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=79 or rand=88 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������ð��������ʨ�ӣ���ս�»��ǲ�����ʨ�ӿ��£�"
                        sql="update �û� set ״̬='��' where ����='" & name1 & "'"
                        conn.execute sql 
                        sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�ڹµ�ð��������','����')" 
                        'sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�ڹµ�ð��������')"
                        conn.execute sql
                        call boot(name1)
                 elseif rand=80 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢�����ϳ˵������ķ����������1000�㣡"
                        sql="update �û� set ����=����-11,����=����+1000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=81 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "������һ��ɫŮ�ӣ���ɫ�Ĵ��𣬶���ʩ���������½�1000��"
                        sql="update �û� set ����=����-1000,����=����-1000,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                 elseif rand=82 or rand=89 then
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "�������Ҳ���ʳ��������ˣ��Ӵ˽�������һ������������"
                        sql="update �û� set ״̬='��' where ����='" & name1 & "'"
                        conn.execute sql 
                        
			sql= "insert into l(b,a,c,e,d) values ('" & name1 & "',now(),'" & aqjh_name & "','�ڹµ�ð��������','����')"
                        'sql="insert into ����(����,ʱ��,����,����) values ('" & name1 & "',now(),'��','�ڹµ�ð��������')"
                        conn.execute sql
                     call boot(name1)
                 
                      
                else
                        huayuan="" & name1 & "�ڹµ���" & sy1 & sh1 & "���㷢����һ���ر��⣬��ûԿ�ף��򲻿��ţ�"
                        sql="update �û� set ����=����-11,����ʱ��=now() where ����='" & name1 & "'"
                        conn.execute sql
                end if
       end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
sub boot(to1)
Application.Lock
onlinelist=Application("aqjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
end if
next
useronlinename=useronlinename&" "
Application.UnLock
end sub
%>