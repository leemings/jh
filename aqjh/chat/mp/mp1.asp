<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����ɴ�ս��<font color=" & saycolor & ">"+menattack(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
'���ɴ�ս(��������)
function menattack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
if Weekday(date())<>1 or (Hour(time())<18 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ɴ�ս����ÿ��������18-22����У�ÿ�ܱ������������Ա�����鿴�Ǽ�����ս����ʣ�����,ս��������߽��ɵ���վ��Ѱ����Ա.��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
	sql="select ����,����,�书,����,����,����,���,�ȼ�,��� from �û� where ����='" & aqjh_name &"'"
	rs.open sql,conn,1,1
       mp=rs("����")
       zhan1=rs("����")
       gong1=rs("����")
       zhan1=rs("����")
       wg=rs("�书")
       jin1=rs("���")
       yx=rs("����")
       shenfen1=trim(rs("���"))
       dengji1=rs("�ȼ�")
       if mp="��" or mp="" then
       Response.Write "<script language=JavaScript>{alert('���������������ˣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
   
     if mp="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô�ܸ��������ɽ�Թ�ˣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
         if zhan1<=1000 then
       Response.Write "<script language=JavaScript>{alert('����ս������ô��,��ô�����ɳ�ս����');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if shenfen1="����" then
          zhishu1=2
       elseif shenfen1="������" then
         zhishu1=1.8
       elseif shenfen1="����"  then
         zhishu1=1.6
       elseif shenfen1="����"  then
          zhishu1=1.4
        elseif shenfen1="����"  then
          zhishu1=1.2
        else
           zhishu1=1
        end if
        rs.close
       sql="SELECT q,f FROM p WHERE  a='" & mp & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	   shidi1=rs("q")
	   baohu1=rs("f")
	   if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('�����������ڽ��ܹٸ�������,���޷�������ս');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="select ����,grade,����,����,���,����,�ȼ�,��� from �û� where ����='" & to1 & "'"
       rs.open sql,conn,1,1
       topai=rs("����")
       fang1=rs("����")
       yin2=rs("���")
       ti=rs("����")
       dengji2=rs("�ȼ�")
       zhan2=rs("����")
       shenfen2=trim(rs("���")) 
      
     
   if topai="��" or topai="" then
       Response.Write "<script language=JavaScript>{alert('�Է������������ɣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 if topai="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է����ǹٸ����㲻�����');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
       if InStr("|" & shidi1 & "|", "|" & topai& "|")=0 and mp<>topai then
       Response.Write "<script language=JavaScript>{alert('�������ɺ���û�к�������ɿ�սѽ,����뿪ս������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
 
       
	 if shenfen2="����" then
          zhishu2=2
       elseif shenfen2="������" then
         zhishu2=1.8
       elseif shenfen2="����"  then
         zhishu2=1.6
       elseif shenfen2="����"  then
          zhishu2=1.4
        elseif shenfen2="����"  then
          zhishu2=1.2
        else
           zhishu2=1
        end if
       rs.close
       sql="SELECT s,h,q,f FROM p WHERE  a='" & topai & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
        shengming=rs("s")
        kunjin=int(rs("h"))
        shidi2=rs("q")
         baohu2=rs("f")
          if baohu2=1 then
       Response.Write "<script language=JavaScript>{alert('�����������ڽ��ܹٸ�������,���޷�������ս');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       if InStr("|" & shidi2 & "|", "|" & mp& "|")=0 and mp<>topai then
       Response.Write "<script language=JavaScript>{alert('�������ɺ���û�к�������ɿ�սѽ,����뿪ս������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
'��ʼ�书
rs.open "select * FROM y WHERE a='" & trim(fn1) & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������:"& mp &" ��û���������书["& fn1 &"]ѽ��');}</script>"
	Response.End
end if
nei=abs(rs("d"))
       if nei<100 then
       Response.Write "<script language=JavaScript>{alert('������Ҳ̫���˰�');}</script>"
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if   
       if zhan1<nei or zhan1<=1000 then
       Response.Write "<script language=JavaScript>{alert('����������㣬С�����ѣ���ȥ������');}</script>"
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if 
 
	if mp=topai then
	    pang=2
	 else 
	    pang=1
	 end if         
       Randomize
	m1 = Int(20* Rnd + 1)
       Select Case m1
          case "1"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>��" & to1 & "ʹ����һ��" & mp & "��("& fn1 &")��ǡ��"&to1&"���Ա�һ�Σ��������һ�٣�����Ŷ!"
          baofa=0
          case"2"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>����" & to1 & "������͵Ϯ�ɹ���һ��" & mp & "��("& fn1 &")�����صĴ���"&to1&"�ı��ϣ�"
          baofa=2
          case"8"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>ͻ�����������һ����һ��" & mp & "��("& fn1 &")���"&to1&"������Ѫ��"
          baofa=3
          case"16"
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>͵Ϯ"&to1&"������"&to1&"���з���������"&to1&"�Ļ��������ˣ������½�-5000��"
          baofa=0
          conn.execute"update �û� set ,����=����-5000 where ����='" & aqjh_name & "'"
          case else
          baofa=1
          e="[<a href=javascript:parent.sw('["&aqjh_name&"]'); target=f2>"&aqjh_name & "</a>]<bgsound src=READONLY/fazhao.mid loop=1>��������"&to1&"ʹ����һ��" & mp & "��("& fn1 &")��"
       end select
          if baofa=0 then
            kills=0 
            else
            kills1=nei*(gong1+1)*baofa+int(jin1/100)*dengji1     
            kills2=pang*zhishu1*kills1/zhishu2
            kills=int(kills2/(fang1+1))
            kills3=int(kills2/100)
          end if
           if kills<=100 then
	      randomize timer
	      kills=int(rnd*99)+1
          end if
	            	conn.execute"update p set s=s-"& kills3 & " where a='" & topai & "'"

	             	conn.execute"update �û� set ����=����-"  & kills & " where ����='" & to1 & "'"
			         if shengming-kills3<=0 then
			              if mp=topai then
			              conn.execute"update �û� set ���=���+10,����=����+1000,����=����+1000,����=����*2,�书=�书*2 where ����='" & aqjh_name & "'"
			              conn.execute"update �û� set ����='��',���='��',����=����*1/2,�书=�书*1/2 where ����='" & topai & "'"
                          conn.execute"delete * from p  where a='" & topai & "'"
                          call boot(to1,"���ɴ�ս�������ߣ�"&aqjh_name&",["&mp&"]"&fn1)
                          conn.execute"insert into ����(����,ʱ��,����,����) values ('" & topai & "',now(),'" &aqjh_name & "','�������ڼ����')"
                         

			              menattack=e&"ɱ��"&to1&"ս����Ϊ<b><font color=red>"&kills&"</font></b>,ɱ��"&topai&"����ս����Ϊ<b><font color=red>"&kills3&"</font></b>,������ʹ�����ɸ���,"& aqjh_name &"��Ϊ�Ե׳������,��������1000,��������1000,�������<b><font color=red>10</font></b>!"
                          else 
			            shidis=replace(shidi1,topai&"|","")
			            conn.execute"update �û� set ���=���+10,����=����+1000,����=����+1000 where ����='" & aqjh_name & "'"
                    	conn.execute"update �û� set ����='��',���='��',����=����*1/2,�书=�书*1/2 where ����='" & topai & "'"
						conn.execute"delete * from p  where a='" & topai & "'"
						conn.execute"update ����  set h=h+'"&kunjin&"',q='"&shidis&"' where a='" & mp & "'"
						call boot(to1,"���ɴ�ս�������ߣ�"&aqjh_name&",["&mp&"]"&fn1)
                        conn.execute"insert into ����(����,ʱ��,����,����) values ('" & topai & "',now(),'" & mp & "','�����Χ�˸���')"
                        menattack=e&"ɱ��"&to1&"ս����Ϊ<b><font color=red>"&kills&"</font></b>,ɱ�˶Է�"&topai&"����ս����Ϊ<b><font color=red>"&kills3&"</font></b>,������ʹ�����ɸ���,"& aqjh_name &"Ϊ���ɳ������,��������1000,��������1000,�������<b><font color=red>10</font></b>!"
						 end if
                             rs.close
                             set rs=nothing	
	                      conn.close
	                      set conn=nothing	
                             exit function
                    end if
                        conn.execute"update �û� set ����=����-" & nei & ",����=����+2 where ����='" & aqjh_name & "'"   
                        menattack=e&"ɱ��"&to1&"ս����Ϊ<b><font color=red>"&kills&"</font></b>,ɱ�˶Է�"&topai&"����ս����Ϊ<b><font color=red>"&kills3&"</font></b>,"& aqjh_name &"Ϊ���ɳ������,��������2!"
                       if zhan2-kills<1000 then
                       menattack=menattack&""&to1&"ս������Ϊ1000����,�Ѿ�ɥʧ��в��,����Ҫ������."
                       end if
                        rs.close
                        set rs=nothing	
	                 conn.close
                       set conn=nothing                              
          
end function
%>