clear all
close all
clc

%========================
%  ���ƂƂȂ�f�[�^
%========================
t(:,1) = 0:0.001:1;
data_a(:,1) = sin(2*pi*1*t);
data_b(:,1) = cos(2*pi*10*t);

%========================
%  �`��
%========================
fig1 = figure(1);
subplot(221);
plot(t,[data_a,data_b]);
xlim([0 0.5]);

legend({'a','b'});
title test
xlabel xtest
ylabel ytest

subplot(222);
plot(t,[data_a,data_b]);
xlim([0 0.2])
legend({'a','b'});

subplot(817);
plot(t,data_b);
xlim([0 0.8])
legend({'b'});

%========================
%  �쐬�����֐�
%  fig1�̕`��͈͓��̃f�[�^��
%  Figure.xlsx�œf���o���B
%========================
exportFigureData2Excel(fig1,'test2');