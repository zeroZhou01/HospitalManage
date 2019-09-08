function varargout = hsi(varargin)
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @hsi_OpeningFcn, ...
                   'gui_OutputFcn',  @hsi_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT





%ҳ�����
function hsi_OpeningFcn(hObject, eventdata, handles, varargin)
handles.output = hObject;
guidata(hObject, handles);



%��ʾ��������
[STATUS,SHEETS] = xlsfinfo('1.xlsx'); % SHEETS ���зֱ�����cell����
sheets=cellfun(@num2str,SHEETS,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
array_all=[];%�����м���ʱ�������ڴ�ȡÿ��ѭ��ȡ��������
for co=1:length(sheets)
    [a b c]=xlsread('1.xlsx',sheets{co});
	yy=cellfun(@num2str,c,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�    
;%��ʾ��ѯ����������   
    array_all{co}=yy;
end
arr_1=[];
arr_2=[];
for cx=1:length(sheets)
    arr_2=cellstr(array_all{cx});
    arr_1=[arr_1;arr_2];
end
%��ʱ��תΪ�·�
yy_str = arr_1(:,1);
 str = [];
    for m=1:length(yy_str)    
       f=cell2mat(yy_str(m));
     %  f(end-9,end) = []; 
       f=f(7:7);
       str = [str;f];  
    end
    cell_str=cellstr(str);%ȡ�����ַ�����ת��ΪԪ������
          arr_1(:,1) = cell_str;%���ת��
   
set(handles.uitable1,'Data',arr_1);%�󶨵�table


%����18�ϰ���datatableҩƷ����
Datas=get(handles.uitable1,'Data');
Data1=Datas(:,5);
a=str2num(char((Data1)));
yaopinshu18=sum(a)
set(handles.ed_18c,'string',yaopinshu18);
%����18�ϰ��������ܴ���
cishu18=size(a,1)
set(handles.ed_18b,'string',cishu18);
%����18�ϰ���ʵ���ܽ��
Data2=Datas(:,7);
b=str2num(char((Data2)));
jine18=sum(b)
set(handles.ed_18a,'string',jine18);

%���18�ϰ���͵���
kedanjia18=jine18/cishu18
set(handles.ed_18d,'string',kedanjia18);

%ҩƷ����
Datas3=get(handles.uitable1,'Data');
raw=Datas3(:,4);

%ת������
zhonglei=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
a=tabulate(zhonglei);
zhongleishu=size(a,1)
set(handles.ed18_zhonglei,'string',zhongleishu);

%�˿�����
raw=Datas3(:,2);
%ת������
gukeshu=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
a=tabulate(gukeshu);
gukezongshu=size(a,1)
set(handles.ed_gukeshu,'string',gukezongshu);

%�˾�����
renjun=jine18/gukezongshu
set(handles.ed_renjun,'string',renjun);

%ÿ��ƽ�����
pingjunjine=jine18/7
set(handles.ed_pjje,'string',pingjunjine);

%δ֪�ô�
function varargout = hsi_OutputFcn(hObject, eventdata, handles) 
varargout{1} = handles.output;







% --- �������ѽ�������ͼ
function pushbutton1_Callback(hObject, eventdata, handles)
figure;
Datas=get(handles.uitable1,'Data');
Data1=Datas(:,7);
a=str2num(char((Data1)));
yaopinshu18=sum(a);
set(handles.ed_18c,'string',yaopinshu18)
plot(a);
title('��������������ͼ');
xlabel('ʱ��');
ylabel('ʵ�ʽ��');



% �������ѽ��ͼ
function pushbutton3_Callback(hObject, eventdata, handles)
%�������;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%ת������
yy=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
%��ѯ1�µ����ѽ��
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%ȡ�����ݲ����ַ����󱣴�
mon1 = mon1(:,7);%������ȡʵ�����ѽ��
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %���
%��ѯ2�µ����ѽ��
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%ȡ�����ݲ����ַ����󱣴�
mon2 = mon2(:,7);%������ȡʵ�����ѽ��
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%���
%��ѯ3�µ����ѽ��
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%ȡ�����ݲ����ַ����󱣴�
mon3 = mon3(:,7);%������ȡʵ�����ѽ��
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%���
%��ѯ4�µ����ѽ��
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%ȡ�����ݲ����ַ����󱣴�
mon4 = mon4(:,7);%������ȡʵ�����ѽ��
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%���

%��ѯ5�µ����ѽ��
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%ȡ�����ݲ����ַ����󱣴�
mon5 = mon5(:,7);%������ȡʵ�����ѽ��
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%���

%��ѯ6�µ����ѽ��
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%ȡ�����ݲ����ַ����󱣴�
mon6 = mon6(:,7);%������ȡʵ�����ѽ��
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%���

%��ѯ7�µ����ѽ��
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%ȡ�����ݲ����ַ����󱣴�
mon7 = mon7(:,7);%������ȡʵ�����ѽ��
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%���
%��ͼ
figure;
disp('�����ѽ�����');
x = [mon1_jine,mon2_jine,mon3_jine,mon4_jine,mon5_jine,mon6_jine,mon7_jine]
plot(x);
title('�������ѽ��ͼ');
xlabel('�·�');
ylabel('ʵ�����ѽ��');
set(gca,'xticklabel',{'1��','2��','3��','4��','5��','6��','7��'});


% --- �����Ѵ������ͼ
function pushbutton4_Callback(hObject, eventdata, handles)
%�������;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%ת������
yy=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
%��ѯ1�µ����Ѵ���
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu1=size(mon1,1);
%��ѯ2�µ����Ѵ���
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu2=size(mon2,1);
%��ѯ3�µ����Ѵ���
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu3=size(mon3,1);
%��ѯ4�µ����Ѵ���
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu4=size(mon4,1);
%��ѯ5�µ����Ѵ���
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu5=size(mon5,1);

%��ѯ6�µ����Ѵ���
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu6=size(mon6,1);

%��ѯ7�µ����Ѵ���
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu7=size(mon7,1);
%��ͼ
figure;
disp('ÿ�´������');
x = [cishu1,cishu2,cishu3,cishu4,cishu5,cishu6,cishu7]
plot(x);
title('�������Ѵ���ͼ');
xlabel('�·�');
ylabel('����');
set(gca,'xticklabel',{'1��','2��','3��','4��','5��','6��','7��'});



% ---�¿͵�������������ѽ��/������
function pushbutton5_Callback(hObject, eventdata, handles)
%�������;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%ת������
yy=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
%��ѯ1�µ����ѽ����Դ����ÿ͵���
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu1=size(mon1,1);
mon1 = mon1(:,7);%������ȡʵ�����ѽ��
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %���
kedanjia1=mon1_jine/cishu1;%��͵���
%��ѯ2�µ����ѽ����Դ���
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu2=size(mon2,1);
mon2 = mon2(:,7);%������ȡʵ�����ѽ��
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%���
kedanjia2=mon2_jine/cishu2;%��͵���
%��ѯ3�µ����ѽ����Դ���
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu3=size(mon3,1);
mon3 = mon3(:,7);%������ȡʵ�����ѽ��
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%���
kedanjia3=mon3_jine/cishu3;%��͵���
%��ѯ4�µ����ѽ����Դ���
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu4=size(mon4,1);
mon4 = mon4(:,7);%������ȡʵ�����ѽ��
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%���
kedanjia4=mon4_jine/cishu4;%��͵���

%��ѯ5�µ����ѽ����Դ���
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu5=size(mon5,1);
mon5 = mon5(:,7);%������ȡʵ�����ѽ��
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%���
kedanjia5=mon5_jine/cishu5;%��͵���

%��ѯ6�µ����ѽ����Դ���
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu6=size(mon6,1);
mon6 = mon6(:,7);%������ȡʵ�����ѽ��
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%���
kedanjia6=mon6_jine/cishu6;%��͵���
%��ѯ7�µ����ѽ����Դ���
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%ȡ�����ݲ����ַ����󱣴�
cishu7=size(mon7,1);
mon7 = mon7(:,7);%������ȡʵ�����ѽ��
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%���
cishu7=size(mon7,1);
kedanjia7=mon7_jine/cishu7;%��͵���
%��ͼ
figure;
disp('�͵������');
x = [kedanjia1,kedanjia2,kedanjia3,kedanjia4,kedanjia5,kedanjia6,kedanjia7]
plot(x);
title('�¿͵���ͼ');
xlabel('�·�');
ylabel('�͵���');
set(gca,'xticklabel',{'1��','2��','3��','4��','5��','6��','7��'});


% --- ҩƷ��������ͼ
function pushbutton6_Callback(hObject, eventdata, handles)
%�������;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%ת������
yy=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
%��ѯ�����ᰱ�ȵ�ƽƬ(������)
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'�����ᰱ�ȵ�ƽƬ(������)'),yy(:,4)));
mon1 = yy(mon1_row,:);%ȡ�����ݲ����ַ����󱣴�
mon1 = mon1(:,5);%������ȡʵ������
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %���
%��ѯ2����ͨ
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'����ͨ'),yy(:,4)));
mon2 = yy(mon2_row,:);%ȡ�����ݲ����ַ����󱣴�
mon2 = mon2(:,5);%������ȡʵ������
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%���
%��ѯ��ʯ���������Ƭ(�����ֿ�)
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'��ʯ���������Ƭ(�����ֿ�)'),yy(:,4)));
mon3 = yy(mon3_row,:);%ȡ�����ݲ����ַ����󱣴�
mon3 = mon3(:,5);%������ȡʵ������
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%���
%��ѯ������ƽƬ(��ʹ��)
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'������ƽƬ(��ʹ��)'),yy(:,4)));
mon4 = yy(mon4_row,:);%ȡ�����ݲ����ַ����󱣴�
mon4 = mon4(:,5);%������ȡʵ������
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%���

%��ѯ�����ᰱ�ȵ�ƽƬ(���ϲ)
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'�����ᰱ�ȵ�ƽƬ(���ϲ)'),yy(:,4)));
mon5 = yy(mon5_row,:);%ȡ�����ݲ����ַ����󱣴�
mon5 = mon5(:,5);%������ȡʵ������
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%���

%��ѯ'������ѪƽƬ(������ѹƬ)
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'������ѪƽƬ(������ѹƬ)'),yy(:,4)));
mon6 = yy(mon6_row,:);%ȡ�����ݲ����ַ����󱣴�
mon6 = mon6(:,5);%������ȡʵ������
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%���

%��ѯG�����������������Ƭ(�����ֿ�)
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'G�����������������Ƭ(�����ֿ�)'),yy(:,4)));
mon7 = yy(mon7_row,:);%ȡ�����ݲ����ַ����󱣴�
mon7 = mon7(:,5);%������ȡʵ������
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%���
%��ѯ��ɳ̹����(����)
[mon8_row mon8_col]=find(cellfun(@(x) strcmp(x,'��ɳ̹����(����)'),yy(:,4)));
mon8 = yy(mon8_row,:);%ȡ�����ݲ����ַ����󱣴�
mon8 = mon8(:,5);%������ȡʵ������
mon8=str2num(char((mon8)));
mon8_jine =sum(char((mon8)));%���

%��ѯ�����ƽ����Ƭ(������)
[mon9_row mon9_col]=find(cellfun(@(x) strcmp(x,'�����ƽ����Ƭ(������)'),yy(:,4)));
mon9 = yy(mon9_row,:);%ȡ�����ݲ����ַ����󱣴�
mon9 = mon9(:,5);%������ȡʵ������
mon9=str2num(char((mon9)));
mon9_jine =sum(char((mon9)));%���


%��ѯ������
[mon10_row mon10_col]=find(cellfun(@(x) strcmp(x,'������'),yy(:,4)));
mon10 = yy(mon10_row,:);%ȡ�����ݲ����ַ����󱣴�
mon10 = mon10(:,5);%������ȡʵ������
mon10=str2num(char((mon10)));
mon10_jine =sum(char((mon10)));%���

%��ͼ
figure;
disp('ҩƷ�����������');
x = [mon1_jine,mon2_jine,mon3_jine,mon4_jine,mon5_jine,mon6_jine,mon7_jine,mon8_jine,mon9_jine,mon10_jine]
bar(x);
title('ҩƷ�������ͼ');;
ylabel('����');
set(gca,'xticklabel',{'������','����ͨ','�����ֿ�','��ʹ��','���ϲ','������ѹƬ','�����ֿ�','����','������','������'});
 xtl=get(gca,'XTickLabel'); 
 % ��ȡxtick��ֵ
 xt=get(gca,'XTick'); 
% ��ȡytick��ֵ         
yt=get(gca,'YTick');   
% ����text��x����λ����         
xtextp=xt;                   
 % ����text��y����λ����      
 ytextp=(yt(1)-0.2*(yt(2)-yt(1)))*ones(1,length(xt)); 
% rotation��������ת�Ƕȴ�����ʱ����ת����ת�������HorizontalAlignment�������趨��
% ��3������ֵ��left��right��center
 text(xtextp,ytextp,xtl,'HorizontalAlignment','right','rotation',90,'fontsize',12); 
% ȡ��ԭʼticklabel
 set(gca,'xticklabel','');


% ---ҩƷ���۴����ݵ�����cmd
function pushbutton7_Callback(hObject, eventdata, handles)
Datas=get(handles.uitable1,'Data');
raw=Datas(:,4);
%ת������
yy=cellfun(@num2str,raw,'UniformOutput',false);%��cell�е�����ת��Ϊ�ַ�
disp('ҩƷ���۴������');
a=tabulate(yy)


%������edittext��Ӧ�ĺ���
function ed_18a_CreateFcn(hObject, eventdata, handles)
function ed_pjje_CreateFcn(hObject, eventdata, handles)
function ed_renjun_CreateFcn(hObject, eventdata, handles)
function ed_gukeshu_CreateFcn(hObject, eventdata, handles)
function ed18_zhonglei_CreateFcn(hObject, eventdata, handles)
function ed_18c_CreateFcn(hObject, eventdata, handles)
function ed_18d_CreateFcn(hObject, eventdata, handles)
function ed_18b_CreateFcn(hObject, eventdata, handles)









