<%@ LANGUAGE=VBScript codepage ="936"%>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

pet0=8		'�����տ�������,Ĭ��Ϊ8Сʱ
pet1=1		'ÿ��Сʱ�۵ĵ���(��,��,��,����)
pet111=2		'ÿ��Сʱ�۵�����
pet2=10		'ι��ʱ����������
pet3=30		'ι��ʱ�ֻ�����������

pet4=5		'����ʱ���ӵ�:��,��
pet41=5		'����ʱ���ӵ�:����
pet5=2		'����ʱ���ٵ�:����
pet6=15	'����ʱ���ٵ�:����

pet7=10		'���ʱ���ӵ�:����
pet8=2		'���ʱ���ӵ�:����
pet9=15	'���ʱ���ٵ�:����

pet10=3		'ϴ��ʱ���ӵ�:����
pet11=2		'ϴ��ʱ���ӵ�:��,��
pet12=15	'ϴ��ʱ���ٵ�:����
pet121=3		'ϴ��ʱ���ٵ�:����

pet13=8		'��ѵʱ���ӵ�:��,��
pet14=3	'��ѵʱ���ӵ�:����
pet15=10	'��ѵʱ���ӵ�:����
pet16=15	'��ѵʱ���ٵ�:����

pet17=20	'��Ϣʱ���ӵ�:����
pet18=2		'��Ϣʱ���ӵ�:����
pet19=2		'��Ϣʱ���ӵ�:����
pet20=2	'��Ϣʱ���ӵ�:��,��
pet21=1000000	'����ʱ���������
pet22=15		'�������������һ����������15��


set xajh=Server.CreateObject("anjh70.xajh")
call xajh.pet(pet0,pet1,pet111,pet2,pet3,pet4,pet41,pet5,pet6,pet7,pet8,pet9,pet10,pet11,pet12,pet121,pet13,pet14,pet15,pet16,pet17,pet18,pet19,pet20,pet21,pet22)
set xajh=nothing
'call pet()

%>