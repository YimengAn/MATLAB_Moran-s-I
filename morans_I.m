%  程序名称: 莫兰指数计算
%  数据输入：X1=面板数据；W=距离权重
%  结果输出：G_moran_I %第一列为莫兰指数，第二列为Z值，第三列为p值，每一行为一个年份。
clc,clear all;
format short g;
fprintf(2,'Moran Idx：\n')

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%% 数据输入 %%%%%%%%%%%%%%%%%%%%%%%%%
%% 读入数据
sheetNames = sheetnames("pollutiondata.xlsx");
polltdata = readtable('pollutiondata.xlsx',"ReadRowNames",false,"ReadVariableNames",...
        false,"Sheet",sheetNames{1},"Range","B4:AP87"); % 大气污染数据
polltdata = polltdata{:,:}; % table 转 matrix
[T,Ind] = size(polltdata); % T=时间长度;Ind=研究对象个数
sp_weit = readtable('pollutiondata.xlsx',"ReadRowNames",false,"ReadVariableNames",...
        false,"Sheet",sheetNames{2},"Range","F3:AT43"); % 空间权重矩阵
sp_weit = sp_weit{:,:};

%% %%%%%%%%%%%%%%%%%%%%%%%%%%% 莫兰指数计算 %%%%%%%%%%%%%%%%%%%%%%%%%
% 此部分源代码转载自: https://blog.csdn.net/wbj3106/article/details/81984530
%{
输入数据
X1 = [];%放入变量,按列存放，每一列是一年的数据。
W = [];%按行归一化空间矩阵
%}

X1 = polltdata';
W = sp_weit;
[m,n] = size(X1); %共m个样本，n年的数据。
G_moran_I = zeros(n,3);
H = -1/(m-1);%期望
S0 = sum(W(:));
S1 = sum(sum((W+W').^2))/2;
S2 = sum((sum(W,2)+sum(W,1)').^2);
D = (m^2*S1-m*S2+3*S0^2)/S0^2/(m^2-1)-H^2;
Std = sqrt(D); %方差

for q=1:n
    X=X1(:,q);
    x_=mean(X);
    s=var(X,1);
    x1=(X-x_);
    moran_I=x1'*W*x1/m/s;
    Z=(moran_I-H)/Std;
    p=1-normcdf(abs(Z),0,1);
    G_moran_I(q,:)=[moran_I Z p];
end

%  结果输出
disp('G_moran_I:');
disp(G_moran_I); %第一列为莫兰指数，第二列为Z值，第三列为p值，每一行为一个年份。

%% %%%%%%%%%%%%%%%%%%%%%%%%%%% 绘制莫兰指数直方图 %%%%%%%%%%%%%%%%%%%%%%%%%
% Moran_year = [];
% for i = 1:7 % 8年数据
%     moran_year = G_moran_I(12*(i-1)+1:12*i,1);
%     Moran_year = [Moran_year,moran_year];
% end
% Bar = bar(Moran_year',1);
% 
% % 美化图片
% Bar(1).FaceColor = '#ac7c82';
% Bar(2).FaceColor = '#887880';
% Bar(3).FaceColor = '#9f90a2';
% Bar(4).FaceColor = '#bca7b2';
% Bar(5).FaceColor = '#cfbfc7';
% Bar(6).FaceColor = '#f0dbdf';
% Bar(7).FaceColor = '#dbe2e8';
% Bar(8).FaceColor = '#a6b3c3';
% Bar(9).FaceColor = '#849fbb';
% Bar(10).FaceColor = '#6687a8';
% Bar(11).FaceColor = '#60788c';
% Bar(12).FaceColor = '#828f9f';
% 
% set(Bar,'FaceAlpha',1);
% set(gca,'xticklabel',{'2016','2017','2018','2019','2020','2021','2022'},...
%     'FontSize',14);
% legend({'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'},...
%     'Orientation','horizontal','Location','northoutside');
% legend('boxoff');
% xlabel('Date');
% ylabel('Moran Index');


