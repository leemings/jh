<%
'****************ÿ���½�������***************
jl_dd=28      'ÿ������һ����н���(��ǰ��Ϊÿ��28��)
jl_allvalue=16200      '������Ҫ�ﵽһ���Ļ���
jl_yjf_top=10      '�»��ֽ���ǰ������
jl_yjf_jk=20      '�»��ֽ�����ϵ��
jl_lr_jk=20      '���˽�����ϵ��
jl_lr_yl=500000      '���˽�������ϵ��
'****************��ֹע����¼������***************
aqjh_disloginname="���,����,����,����,����,����,�ڿ�,�ٿ�,�й�,������վ,վ��,����,��,��,ү,��,sex,master,grade,�ٸ�,bi,����,��,��,��,��,����,����,��,�侫,��,��ʺ,����,����,����,����,��Ա,����,����,��,����,������,����,��,����,����,���,����,����,����,����,����,����,����,����,����,����,����,غ��,��ȥ,��Ů,����,����,���������,���������,��,�,���������,��Ƥ,��,��,��,��,��ͷ,��,����,����,��,�ظ�,�P,��,�H,����,��,��,����,����,����,����,����,����,ʮ����,��ү,������,���ϰ�,������,������,���,�Ҷ�,����,�Ҹ�,�Ҳ�,��,��,��,��,��,��ĸ,����,�Ϲ�,����,��,ĸ,��,��ľ,��Ĺ,��,��,����,����,����,���,����,���,����,�ϰ�,�ϴ�,��,��,�Ҹ�,ȥ��,����,���׹�,������,����,����"      '��ֹ��¼������
'****************�����ҽ�ֹ���໰***************
aqjh_badword="'��','�Ҹ�','����','ͱ��','ͱ����','��Ů','��Ѽ','��Ѽ','����Ա','����','����ĸ','��ĸ','��ľ','��','��','����','��','����','�Ҳ�','������','����','��','���','����','���','����','�','����','����','��','����','����','����','����','����','����','غ��','��Ƥ','��ͷ','��','�P','��','�H','����','��','��','����','����','����','����','��Ů','���','ʺ','http','www','com','cn','org','net','://','WWW','HTTP','Http','HTTp','HttP','hTTP','htTp','httP','COM','CN','update','UPDATE','grade','allvalue','set','SET','Set','sEt','seT','add','Add','AdD','ADd','ADD','table','ORG','NET','Www','wWw','WWw','wwW','wWW','Net','NEt','nET','neT','Com','COm','cOM','Org','ORg','oRG','orG','coM','fuck','gan','��','����qq','����QQ','���ң��','���ңѣ�','����qq','����QQ','���ң��','���ңѣ�','��QQ','��qq','�ӣ��','�ӣѣ�','Q Q','q q','�� ��','�ѣ�','��','��','��','��','��','��','��','��','��','��','��','��','��','��','��'"      '�����ҽ�ֹ���໰
'****************������λ��chat/picwords�ļ�����***************
aqjh_Zshow="��|��|��|��|��|��|��|��|��|��|��|ס|��|��|ֻ|ֹ|ֱ|֧|��|��|��|��|վ|��|��|��|Զ|ԭ|��|��|��|��|ӵ|Ӱ|��|��|��|��|��|��|Ҳ|Ҫ|��|��|��|��|��|ѩ|ѧ|ѡ|��|��|��|��|��|��|��|л|д|Ц|��|��|��|��|��|��|��|��|��|��|��|��|ϵ|Ϸ|ϰ|��|��|��|�|��|��|ι|Ϊ|��|��|��|��|��|ͼ|ͷ|ͬ|ͨ|��|��|��|̫|��|��|��|��|��|˵|˯|ˮ|˭|˫|��|��|��|��|��|ʿ|ʱ|ʮ|��|��|��|��|ɽ|ɵ|ɫ|��|��|��|��|Ȼ|ȷ|ȥ|��|��|Ǯ|ǰ|ǧ|��|��|��|��|��|��|��|Ƽ|Ƶ|Ư|Ƭ|��|��|��|ż|Ŷ|ţ|��|��|��|��|��|��|��|��|��|ľ|��|��|��|��|��|ü|è|æ|��|��|��|·|¼|�|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|��|ô|��|��|��|ˬ|��|��|��|��|��|С|��|��|��|��|��|0|1|2|3|4|5|6|7|8|9|A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z|��|"      '����

%>