
<%
Response.buffer=true
reno=trim(cstr(application("reno")))      
charp=request("charp")  
ddwho=trim(request("ddwho"))
    if  ddwho="1" or ddwho="2" or ddwho="3" or ddwho="4"   then  
    application("dddwho")=ddwho
    end if
    if  charp="stopp" or  charp="addmoney" or charp="nomoney" or charp="myhouse" or   charp="pmoney"   then application("charp")=charp	

if charp="stopp" and ddwho<>"" then session("stopp")=""
if charp="nomoney" and ddwho<>""  then session("nomoney")=""
if charp="addmoney" and ddwho<>""  then session("addmoney")=""
if charp="pmoney" and ddwho<>""  then session("pmoney")=""
if charp="myhouse"  and ddwho<>"" then session("myhouse")=""



 if reno="1" then bname="С����"
 if reno="2" then bname="��������"
 if reno="3" then bname="Ǯ����"
 if reno="4" then bname="ɳ¡��˹"
  if application("dddwho")="1" then toname="С����"
  if application("dddwho")="2" then toname="��������"
  if application("dddwho") ="3" then toname="Ǯ����"
  if application("dddwho")  ="4"then  toname="ɳ¡��˹"
 
 if   session("players")=1 then sname="С����"
 if   session("players")=2 then sname="��������"
 if  session("players")=3  then sname="Ǯ����"
 if   session("players")=4 then sname="ɳ¡��˹"




saymsg=left(trim(request("saymsg")),23)
if saymsg<> "" then 
saymsg=trim(saymsg)
saymsg="<font color=#339933>"&sname&"˵<b>:</b>"&saymsg&" </font>"
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")=saymsg
application.unlock
end if
'11
if application("reno")=session("players") then 
session("waityou")=1

 if  trim(application("dddwho"))=trim(cstr(session("players"))) then 
  charp=trim(application("charp"))
if charp="stopp" then session("zjstopp")=1  
if charp="nomoney" then session("zjnomoney")=1  
application("dddwho")=""
application("charp")=""
end if



if application("charp")<>""    and charp<>""  then   'ע������charp
if application("charp")="stopp" then charpname="ͣ����"
if application("charp")="nomoney" then charpname="��˰��"
if application("charp")="addmoney" then charpname="��˰��"
if application("charp")="myhouse" then charpname="���ؿ�"
if application("charp")="pmoney" then charpname="������"

IF application("charp")<>"myhouse"   then 

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
if bname=toname then toname="�Լ�"
application("msgg1")=" <font color=red>"&bname&"��"&toname&"ʹ��"&charpname&"</font>"

if application("charp")="addmoney" then 
application("msgg1")=" <font color=red>"&bname&"��"&toname&"��˰</font>��˰"&fix((application("money"&ddwho)*0.4))&"Ԫ���Լ�,���޺�Ʋ�����~~��~~~��"
end if

if application("charp")="pmoney" then 
application("msgg1")=" <font color=red>"&bname&"��"&toname&"ʹ��"&charpname&"</font>,���������ֽ�"
end if

application.unlock

session("stopp")=""
end if
	END IF

myhome=1  '���صķ��Ӻ�
dj=1   ' '���صķ��ӵȼ���ʼֵ
IF application("dddwho")<>"" AND application("charp")="myhouse" THEN
for i=1 to 40 
     if    mid(trim(application("house"&I)),1,1)=trim(application("dddwho")) then 
	if   mid(trim(application("house"&i)),2,1)>dj then
	     dj=mid(trim(application("house"&i)),2,1)
	     myhome=i
  	   end if
	end if
next 

myhomez=""&trim(reno)&""&trim(dj)&""
application("house"&myhome)=trim(myhomez)

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")=" <font color=red>"&bname&"��"&toname&"ʹ��"&charpname&"</font>,��ú�լ"&myhome&"~"&dj&"����~~��~~~��"
application.unlock
application("dddwho")=""
application("charp")=""
session("myhouse")=""
END IF

 if    application("charp")="pmoney" then 
ddwho=trim(application("dddwho"))
allmoney=application("money"&reno)+application("money"&ddwho)
allmoney=fix(allmoney/2)
application("money"&reno)=allmoney
application("money"&ddwho)=allmoney
'response.write application("money"&reno)
'response.write "<br>"
'response.write application("money"&ddwho)

'response.end
application("charp")=""
end if




if application("charp")="addmoney"    then 
ddwho=trim(application("dddwho"))
application("money"&reno)=application("money"&reno)+fix((application("money"&ddwho)*0.4))
application("money"&ddwho)=application("money"&ddwho)-fix((application("money"&ddwho)*0.4))
application.unlock
application("dddwho")=""
application("charp")=""
end if 
if session("zjstopp")<>""   then  
	 application.lock
	   for i=10 to 2 step-1
	   j=i-1
	   application("msgg"&i)=application("msgg"&j)
	   next
application("msgg1")="ͣ������Ч,"&bname&"ԭ��ͣ��"&session("zjstopp")&"�غ�"
application.unlock
application("reno")=application("reno")+1
application("dddwho")=""
application("charp")=""
session("zjstopp")=session("zjstopp")+1
if session("zjstopp")>3 then session("zjstopp")=""     '  ͣ������غ�
response.redirect "nan.asp"

END IF  


randomize
weiz=rnd
weiz=midb(weiz,5,1)  '���������
if weiz>6 then weiz=6 '���ü���
if weiz=0 then weiz=1 '���ü���
session("weiz")=weiz

if session("players")=1  then 
application("play1")=clng(application("play1"))
application("play1")=application("play1")+weiz 'λ���ۼ�
if  application("play1")>=40 then   application("play1")=application("play1")-39 
'       application("play1")=17
       application("play1")=cstr(application("play1"))

end if

if session("players")=2 then 
	application("play2")=clng(application("play2"))
                  application("play2")=application("play2")+weiz 'ÿ����ҵĲ�ͬ��λ��

                  if  application("play2")>=40 then application("play2")=application("play2")-39 '������ֻ��40��ʱ�����Ը��ݵ�ͼ�Ĵ�С����
'       application("play2")=7

'22
	application("play2")=cstr(application("play2"))
                 end if

if session("players")=3 then 
   application("play3")=clng(application("play3"))
    application("play3")=application("play3")+weiz 'ÿ����ҵĲ�ͬ��λ��
    if  application("play3")>40 then  application("play3")= application("play3")-39 '������ֻ��40��ʱ�����Ը��ݵ�ͼ�Ĵ�С����
'application("play3")=17

   application("play3")=cstr(application("play3"))
end if


if session("players")=4 then 
   application("play4")=clng(application("play4"))
    application("play4")=application("play4")+weiz 'ÿ����ҵĲ�ͬ��λ��
    if  application("play4")>40 then  application("play4")= application("play4")-39 '������ֻ��40��ʱ�����Ը��ݵ�ͼ�Ĵ�С����

   application("play4")=cstr(application("play4"))
end if


if application("play"&reno)="9" or application("play"&reno)="28" then 
randomize
nohouse=rnd
nohouse=midb(nohouse,5,1)  '���������
response.write nohouse
response.write "<br>"
if nohouse >5 then 
nohouse=2
else
nohouse=1
end if
response.write nohouse
'response.end
if nohouse=1 then  
 	if session("stopp")="" then session("stopp")="stopp"
 	if session("addmoney")="" then session("addmoney")="addmoney"
 	if session("nomoney")="" then session("nomoney")="nomoney"
 	if session("myhouse")="" then session("myhouse")="myhouse"
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")="ͻ���¼�λ��"&application("play"&reno)&" <font color=red>"&bname&"������ſ�Ƭ</font>"
application.unlock
end if

if nohouse=2 then  
dj=1
for i=1 to 40 
     if    mid(trim(application("house"&I)),1,1)=trim(cstr(session("players"))) then 
	if   mid(trim(application("house"&i)),2,1)>dj then
	     dj=mid(trim(application("house"&i)),2,1)
	     myhome=i  
  	   end if
	end if
next 

'myhomez=""&trim(reno)&""&trim(dj)&""  '���ӵ�ֵ,��ǰ��ĳ���
application("house"&myhome)=""

application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
application("msgg1")="ͻ���¼�λ��"&application("play"&reno)&" <font color=red>û��"&bname&"�ķ���"&myhome&"~"&dj&"��</font>"
application.unlock






end if
end if






    reno=trim(cstr(application("reno")))       '���ڵ����,��Ϊ������reno��������һ�����
if  application("play"&reno) ="4"  or  application("play"&reno) ="9" or   application("play"&reno) ="28"    or   application("play"&reno) ="14"  or application("play"&reno) ="20"  or application("play"&reno) =25  or application("play"&reno) ="27"   or application("play"&reno) ="34"  or application("play"&reno) ="38" then 
 application("reno")=application("reno")+1
response.redirect "nan.asp"
end if 

                   strno=trim(cstr(application("play"&reno)))   '��N�������ַ���strno ��ʾ��N������������ڵ�λ��=���ӱ��


                         if application("house"&strno)=""  then
                          application("house"&strno)=cstr(session("players")) '���ǵ�һ����������
application.lock
   for i=10 to 2 step-1
    j=i-1
    application("msgg"&i)=application("msgg"&j)
    next
   application("msgg1")=" "&bname&"<font color='red'>����</font>�յ�"&strno&"�ɹ�����400Ԫ</font>" 
   application.unlock
	       'application("house"&strno)=" "&application("house"&strno)&"1  "
	       application("money"&reno)= application("money"&reno)-400
                   	     END IF

whohouse=mid(trim(application("house"&strno)),1,1)
 if whohouse="1" then toname="С����"
 if whohouse="2" then toname="��������"
 if whohouse="3" then toname="Ǯ����"
 if whohouse="4" then toname="ɳ¡��˹"


houseTop=mid(trim(application("house"&strno)),2,1)
         if  trim(whohouse)=trim(cstr(session("players")))     then 
IF application("house"&strno)<>"" THEN
application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next

	   Select Case  houseTop
	       case ""       
	       application("house"&strno)=" "&application("house"&strno)&"1  "
	      case  "1"       'Ϊ1��ʱ,��Ϊ2��
application("msgg1")=" "&bname&"������"&strno&"<font color='red'>��Ϊ~2��</font>����800Ԫ</font>" 

	       application("money"&reno)= application("money"&reno)-800
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"2  "
	     case  "2"
application("msgg1")=" "&bname&"������"&strno&"<font color='red'>��Ϊ~3��</font>����1600Ԫ</font>" 
                     application("money"&reno)= application("money"&reno)-1600
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"3  "
	     case  "3"     '��Ϊ4��
application("msgg1")=" "&bname&"������"&strno&"<font color='red'>��Ϊ~4��</font>����3200Ԫ</font>" 
	       application("money"&reno)= application("money"&reno)-3200
	       application("house"&strno)=left(trim(application("house"&strno)),1)
                       application("house"&strno)=" "&application("house"&strno)&"4  "
               End Select 
        end if	   
application.unlock
END IF
IF application("house"&strno)<>"" THEN
for st = 1 to 4
         if  trim(whohouse)<>trim(cstr(session("players")))     then    'Ҫ��Ǯ��

 	   ns=cstr(st)
                   	if whohouse= ns  then 	'Ǯ����ns�����
   application.lock
  for i=10 to 2 step-1
   j=i-1
    application("msgg"&i)=application("msgg"&j)
next
	IF    session("zjnomoney")  ="" THEN 
	Select Case  houseTop
	       case "1"      
	  application("money1")=application("money"&ns)+800
	  application("money"&reno)=application("money"&reno)-800
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"�ڷ���"&trim(application("houseno"))&"~1����800Ԫ<font color='red'>��"&toname&"</font>" 

	       case "2"      
	  application("money1")=application("money"&ns)+1600
	  application("money"&reno)=application("money"&reno)-1600
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"�ڷ���"&trim(application("houseno"))&"~2����1600Ԫ<font color='red'>��"&toname&"</font>" 

	       case "3"      
	  application("money1")=application("money"&ns)+3200
	  application("money"&reno)=application("money"&reno)-3200
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"�ڷ���"&trim(application("houseno"))&"~3����3200Ԫ<font color='red'>��"&toname&"</font>" 

	       case "4"      
	  application("money1")=application("money"&ns)+6400
	  application("money"&reno)=application("money"&reno)-6400
                   	application("houseno")=strno
	application("msgg1")=" "&bname&"�ڷ���"&trim(application("houseno"))&"~4����6400Ԫ<font color='red'>��"&toname&"</font>" 
end select
	END IF
'session("zjnomoney")=4
if  session("zjnomoney")<>"" then
	application("msgg1")=" "&bname&"�ڷ���"&trim(application("play"&reno))&"<font color='red'>��˰����Ч"&session("zjnomoney")&"��</font>,���ø�Ǯ��"&toname&"" 
	   session("zjnomoney")=session("zjnomoney")+1
	if session("zjnomoney")>4 then session("zjnomoney")=""
end if
end if
application.unlock
END IF

NEXT	
	END IF
application("reno")=application("reno")+1
end if
response.redirect "nan.asp"

%>