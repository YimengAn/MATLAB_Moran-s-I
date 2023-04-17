%  程序名称: 局部莫兰指数&莫兰散点图
%  数据输入：Y=指标数据；W=标准化距离权重
%  注意: 该程序运行需要用到'morans_I.m'的计算结果（将'morans_I.m'中'绘制莫兰指数直方图'部分标注%）
%  源代码转载自bilibili科研虾：
%  https://www.bilibili.com/video/BV1Qg411f7VC/?spm_id_from=...
%  333.880.my_history.page.click&vd_source=7795a379751aeb47cc58b20dd16743cf


%% %%%%%%%%%%%%%%%%%%%%%%%%%%% 输入数据 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%{
需要输入Y,W,name
Y = [] % 指标数据
W = [] % 权重矩阵(标准化)
name = {}; % .cell格式
%}
sheetNames = sheetnames("pollutiondata.xlsx");
Y = readtable('pollutiondata.xlsx',"ReadRowNames",false,"ReadVariableNames",...
        false,"Sheet",sheetNames{1},"Range","B96:AP96"); 
Y = Y{:,:}';
name = readcell('pollutiondata.xlsx',"Sheet",sheetNames{2},"Range","E3:E43");
W = sp_weit;

%% %%%%%%%%%%%%%%%%%%%%%%%%%  计算局部莫兰指数  %%%%%%%%%%%%%%%%%%%%%%%%%%%
m = length(Y(:));
YD = Y-mean(Y);
W2 = W*YD;
I = YD.*W2; %局部莫兰指数

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%  绘制莫兰散点图  %%%%%%%%%%%%%%%%%%%%%%%%%%%
scatter(YD,W2,'ro');
grid on 
axis square
% xlabel('PM2.5'),ylabel('lagged PM2.5');
hold on 

% p1=ceil(max(YD)*10)/10;
% p2=ceil((-10)*min(YD))/(-10);
% p3=ceil(max(W2)*10)/10;
% p4=ceil((-10)*min(W2))/(-10);
% x = [p2,p1];y = [0,0];plot(x,y);
% x = [0,0];y = [p4,p3];plot(x,y);
x = [-40,40];y = [0,0];plot(x,y);
x = [0,0];y = [-10,15];plot(x,y);
lsl = lsline;
lsl.Color = '#7E2F8E';
% legend([lsl],{'Morans I'})

id=name';
for i = (1:m)
    text(YD(i),W2(i),id{i});
end
hold off
% results = ols(W2,YD);
% vnames=strvcat('W2','莫兰指数');
% prt(results,vnames)


%% 显著性检验
% H=[];
% for i=1:m
%     [h,p] = ztest(I(i),mean(YD),std(YD))
%     H = [H,h]
% end


