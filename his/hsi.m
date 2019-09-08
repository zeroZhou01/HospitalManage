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





%页面加载
function hsi_OpeningFcn(hObject, eventdata, handles, varargin)
handles.output = hObject;
guidata(hObject, handles);



%显示所有数据
[STATUS,SHEETS] = xlsfinfo('1.xlsx'); % SHEETS 所有分表名称cell类型
sheets=cellfun(@num2str,SHEETS,'UniformOutput',false);%将cell中的数字转换为字符
array_all=[];%建立中间临时量，用于存取每次循环取出的数据
for co=1:length(sheets)
    [a b c]=xlsread('1.xlsx',sheets{co});
	yy=cellfun(@num2str,c,'UniformOutput',false);%将cell中的数字转换为字符    
;%显示查询到的数据行   
    array_all{co}=yy;
end
arr_1=[];
arr_2=[];
for cx=1:length(sheets)
    arr_2=cellstr(array_all{cx});
    arr_1=[arr_1;arr_2];
end
%将时间转为月份
yy_str = arr_1(:,1);
 str = [];
    for m=1:length(yy_str)    
       f=cell2mat(yy_str(m));
     %  f(end-9,end) = []; 
       f=f(7:7);
       str = [str;f];  
    end
    cell_str=cellstr(str);%取出的字符矩阵转化为元胞矩阵
          arr_1(:,1) = cell_str;%完成转化
   
set(handles.uitable1,'Data',arr_1);%绑定到table


%计算18上半年datatable药品总数
Datas=get(handles.uitable1,'Data');
Data1=Datas(:,5);
a=str2num(char((Data1)));
yaopinshu18=sum(a)
set(handles.ed_18c,'string',yaopinshu18);
%计算18上半年销售总次数
cishu18=size(a,1)
set(handles.ed_18b,'string',cishu18);
%计算18上半年实收总金额
Data2=Datas(:,7);
b=str2num(char((Data2)));
jine18=sum(b)
set(handles.ed_18a,'string',jine18);

%金额18上半年客单价
kedanjia18=jine18/cishu18
set(handles.ed_18d,'string',kedanjia18);

%药品种类
Datas3=get(handles.uitable1,'Data');
raw=Datas3(:,4);

%转换数据
zhonglei=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
a=tabulate(zhonglei);
zhongleishu=size(a,1)
set(handles.ed18_zhonglei,'string',zhongleishu);

%顾客总数
raw=Datas3(:,2);
%转换数据
gukeshu=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
a=tabulate(gukeshu);
gukezongshu=size(a,1)
set(handles.ed_gukeshu,'string',gukezongshu);

%人均消费
renjun=jine18/gukezongshu
set(handles.ed_renjun,'string',renjun);

%每月平均金额
pingjunjine=jine18/7
set(handles.ed_pjje,'string',pingjunjine);

%未知用处
function varargout = hsi_OutputFcn(hObject, eventdata, handles) 
varargout{1} = handles.output;







% --- 按天消费金额情况作图
function pushbutton1_Callback(hObject, eventdata, handles)
figure;
Datas=get(handles.uitable1,'Data');
Data1=Datas(:,7);
a=str2num(char((Data1)));
yaopinshu18=sum(a);
set(handles.ed_18c,'string',yaopinshu18)
plot(a);
title('按天消费情况金额图');
xlabel('时间');
ylabel('实际金额');



% 按月消费金额图
function pushbutton3_Callback(hObject, eventdata, handles)
%获得数据;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%转换数据
yy=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
%查询1月的消费金额
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%取出数据并以字符矩阵保存
mon1 = mon1(:,7);%单独获取实际消费金额
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %求和
%查询2月的消费金额
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%取出数据并以字符矩阵保存
mon2 = mon2(:,7);%单独获取实际消费金额
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%求和
%查询3月的消费金额
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%取出数据并以字符矩阵保存
mon3 = mon3(:,7);%单独获取实际消费金额
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%求和
%查询4月的消费金额
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%取出数据并以字符矩阵保存
mon4 = mon4(:,7);%单独获取实际消费金额
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%求和

%查询5月的消费金额
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%取出数据并以字符矩阵保存
mon5 = mon5(:,7);%单独获取实际消费金额
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%求和

%查询6月的消费金额
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%取出数据并以字符矩阵保存
mon6 = mon6(:,7);%单独获取实际消费金额
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%求和

%查询7月的消费金额
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%取出数据并以字符矩阵保存
mon7 = mon7(:,7);%单独获取实际消费金额
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%求和
%绘图
figure;
disp('月消费金额情况');
x = [mon1_jine,mon2_jine,mon3_jine,mon4_jine,mon5_jine,mon6_jine,mon7_jine]
plot(x);
title('按月消费金额图');
xlabel('月份');
ylabel('实际消费金额');
set(gca,'xticklabel',{'1月','2月','3月','4月','5月','6月','7月'});


% --- 月消费次数情况图
function pushbutton4_Callback(hObject, eventdata, handles)
%获得数据;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%转换数据
yy=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
%查询1月的消费次数
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%取出数据并以字符矩阵保存
cishu1=size(mon1,1);
%查询2月的消费次数
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%取出数据并以字符矩阵保存
cishu2=size(mon2,1);
%查询3月的消费次数
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%取出数据并以字符矩阵保存
cishu3=size(mon3,1);
%查询4月的消费次数
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%取出数据并以字符矩阵保存
cishu4=size(mon4,1);
%查询5月的消费次数
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%取出数据并以字符矩阵保存
cishu5=size(mon5,1);

%查询6月的消费次数
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%取出数据并以字符矩阵保存
cishu6=size(mon6,1);

%查询7月的消费次数
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%取出数据并以字符矩阵保存
cishu7=size(mon7,1);
%绘图
figure;
disp('每月次数情况');
x = [cishu1,cishu2,cishu3,cishu4,cishu5,cishu6,cishu7]
plot(x);
title('按月消费次数图');
xlabel('月份');
ylabel('次数');
set(gca,'xticklabel',{'1月','2月','3月','4月','5月','6月','7月'});



% ---月客单价情况（月消费金额/次数）
function pushbutton5_Callback(hObject, eventdata, handles)
%获得数据;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%转换数据
yy=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
%查询1月的消费金额除以次数得客单价
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'1'),yy(:,1)));
mon1 = yy(mon1_row,:);%取出数据并以字符矩阵保存
cishu1=size(mon1,1);
mon1 = mon1(:,7);%单独获取实际消费金额
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %求和
kedanjia1=mon1_jine/cishu1;%求客单价
%查询2月的消费金额除以次数
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'2'),yy(:,1)));
mon2 = yy(mon2_row,:);%取出数据并以字符矩阵保存
cishu2=size(mon2,1);
mon2 = mon2(:,7);%单独获取实际消费金额
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%求和
kedanjia2=mon2_jine/cishu2;%求客单价
%查询3月的消费金额除以次数
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'3'),yy(:,1)));
mon3 = yy(mon3_row,:);%取出数据并以字符矩阵保存
cishu3=size(mon3,1);
mon3 = mon3(:,7);%单独获取实际消费金额
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%求和
kedanjia3=mon3_jine/cishu3;%求客单价
%查询4月的消费金额除以次数
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'4'),yy(:,1)));
mon4 = yy(mon4_row,:);%取出数据并以字符矩阵保存
cishu4=size(mon4,1);
mon4 = mon4(:,7);%单独获取实际消费金额
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%求和
kedanjia4=mon4_jine/cishu4;%求客单价

%查询5月的消费金额除以次数
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'5'),yy(:,1)));
mon5 = yy(mon5_row,:);%取出数据并以字符矩阵保存
cishu5=size(mon5,1);
mon5 = mon5(:,7);%单独获取实际消费金额
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%求和
kedanjia5=mon5_jine/cishu5;%求客单价

%查询6月的消费金额除以次数
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'6'),yy(:,1)));
mon6 = yy(mon6_row,:);%取出数据并以字符矩阵保存
cishu6=size(mon6,1);
mon6 = mon6(:,7);%单独获取实际消费金额
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%求和
kedanjia6=mon6_jine/cishu6;%求客单价
%查询7月的消费金额除以次数
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'7'),yy(:,1)));
mon7 = yy(mon7_row,:);%取出数据并以字符矩阵保存
cishu7=size(mon7,1);
mon7 = mon7(:,7);%单独获取实际消费金额
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%求和
cishu7=size(mon7,1);
kedanjia7=mon7_jine/cishu7;%求客单价
%绘图
figure;
disp('客单价情况');
x = [kedanjia1,kedanjia2,kedanjia3,kedanjia4,kedanjia5,kedanjia6,kedanjia7]
plot(x);
title('月客单价图');
xlabel('月份');
ylabel('客单价');
set(gca,'xticklabel',{'1月','2月','3月','4月','5月','6月','7月'});


% --- 药品销售数量图
function pushbutton6_Callback(hObject, eventdata, handles)
%获得数据;
Datas=get(handles.uitable1,'Data');
raw=Datas;
%转换数据
yy=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
%查询苯磺酸氨氯地平片(安内真)
[mon1_row mon1_col]=find(cellfun(@(x) strcmp(x,'苯磺酸氨氯地平片(安内真)'),yy(:,4)));
mon1 = yy(mon1_row,:);%取出数据并以字符矩阵保存
mon1 = mon1(:,5);%单独获取实际数量
mon1=str2num(char((mon1)));
mon1_jine =sum(num2str((char((mon1))))); %求和
%查询2开博通
[mon2_row mon2_col]=find(cellfun(@(x) strcmp(x,'开博通'),yy(:,4)));
mon2 = yy(mon2_row,:);%取出数据并以字符矩阵保存
mon2 = mon2(:,5);%单独获取实际数量
mon2=str2num(char((mon2)));
mon2_jine =sum(char((mon2)));%求和
%查询酒石酸美托洛尔片(倍他乐克)
[mon3_row mon3_col]=find(cellfun(@(x) strcmp(x,'酒石酸美托洛尔片(倍他乐克)'),yy(:,4)));
mon3 = yy(mon3_row,:);%取出数据并以字符矩阵保存
mon3 = mon3(:,5);%单独获取实际数量
mon3=str2num(char((mon3)));
mon3_jine =sum(char((mon3)));%求和
%查询硝苯地平片(心痛定)
[mon4_row mon4_col]=find(cellfun(@(x) strcmp(x,'硝苯地平片(心痛定)'),yy(:,4)));
mon4 = yy(mon4_row,:);%取出数据并以字符矩阵保存
mon4 = mon4(:,5);%单独获取实际数量
mon4=str2num(char((mon4)));
mon4_jine =sum(char((mon4)));%求和

%查询苯磺酸氨氯地平片(络活喜)
[mon5_row mon5_col]=find(cellfun(@(x) strcmp(x,'苯磺酸氨氯地平片(络活喜)'),yy(:,4)));
mon5 = yy(mon5_row,:);%取出数据并以字符矩阵保存
mon5 = mon5(:,5);%单独获取实际数量
mon5=str2num(char((mon5)));
mon5_jine =sum(char((mon5)));%求和

%查询'复方利血平片(复方降压片)
[mon6_row mon6_col]=find(cellfun(@(x) strcmp(x,'复方利血平片(复方降压片)'),yy(:,4)));
mon6 = yy(mon6_row,:);%取出数据并以字符矩阵保存
mon6 = mon6(:,5);%单独获取实际数量
mon6=str2num(char((mon6)));
mon6_jine =sum(char((mon6)));%求和

%查询G琥珀酸美托洛尔缓释片(倍他乐克)
[mon7_row mon7_col]=find(cellfun(@(x) strcmp(x,'G琥珀酸美托洛尔缓释片(倍他乐克)'),yy(:,4)));
mon7 = yy(mon7_row,:);%取出数据并以字符矩阵保存
mon7 = mon7(:,5);%单独获取实际数量
mon7=str2num(char((mon7)));
mon7_jine =sum(char((mon7)));%求和
%查询缬沙坦胶囊(代文)
[mon8_row mon8_col]=find(cellfun(@(x) strcmp(x,'缬沙坦胶囊(代文)'),yy(:,4)));
mon8 = yy(mon8_row,:);%取出数据并以字符矩阵保存
mon8 = mon8(:,5);%单独获取实际数量
mon8=str2num(char((mon8)));
mon8_jine =sum(char((mon8)));%求和

%查询非洛地平缓释片(波依定)
[mon9_row mon9_col]=find(cellfun(@(x) strcmp(x,'非洛地平缓释片(波依定)'),yy(:,4)));
mon9 = yy(mon9_row,:);%取出数据并以字符矩阵保存
mon9 = mon9(:,5);%单独获取实际数量
mon9=str2num(char((mon9)));
mon9_jine =sum(char((mon9)));%求和


%查询高特灵
[mon10_row mon10_col]=find(cellfun(@(x) strcmp(x,'高特灵'),yy(:,4)));
mon10 = yy(mon10_row,:);%取出数据并以字符矩阵保存
mon10 = mon10(:,5);%单独获取实际数量
mon10=str2num(char((mon10)));
mon10_jine =sum(char((mon10)));%求和

%绘图
figure;
disp('药品销售数量情况');
x = [mon1_jine,mon2_jine,mon3_jine,mon4_jine,mon5_jine,mon6_jine,mon7_jine,mon8_jine,mon9_jine,mon10_jine]
bar(x);
title('药品销售情况图');;
ylabel('数量');
set(gca,'xticklabel',{'安内真','开博通','倍他乐克','心痛定','络活喜','复方降压片','倍他乐克','代文','波依定','高特灵'});
 xtl=get(gca,'XTickLabel'); 
 % 获取xtick的值
 xt=get(gca,'XTick'); 
% 获取ytick的值         
yt=get(gca,'YTick');   
% 设置text的x坐标位置们         
xtextp=xt;                   
 % 设置text的y坐标位置们      
 ytextp=(yt(1)-0.2*(yt(2)-yt(1)))*ones(1,length(xt)); 
% rotation，正的旋转角度代表逆时针旋转，旋转轴可以由HorizontalAlignment属性来设定，
% 有3个属性值：left，right，center
 text(xtextp,ytextp,xtl,'HorizontalAlignment','right','rotation',90,'fontsize',12); 
% 取消原始ticklabel
 set(gca,'xticklabel','');


% ---药品销售次数据导出到cmd
function pushbutton7_Callback(hObject, eventdata, handles)
Datas=get(handles.uitable1,'Data');
raw=Datas(:,4);
%转换数据
yy=cellfun(@num2str,raw,'UniformOutput',false);%将cell中的数字转换为字符
disp('药品销售次数情况');
a=tabulate(yy)


%下列是edittext相应的函数
function ed_18a_CreateFcn(hObject, eventdata, handles)
function ed_pjje_CreateFcn(hObject, eventdata, handles)
function ed_renjun_CreateFcn(hObject, eventdata, handles)
function ed_gukeshu_CreateFcn(hObject, eventdata, handles)
function ed18_zhonglei_CreateFcn(hObject, eventdata, handles)
function ed_18c_CreateFcn(hObject, eventdata, handles)
function ed_18d_CreateFcn(hObject, eventdata, handles)
function ed_18b_CreateFcn(hObject, eventdata, handles)









