function [] = exportFigureData2Excel(fig,varargin)
%EXPORTFIGUREDATA2EXCEL Figure�̎��n��f�[�^��Excel�ɏ����o���B
% 
%   Figure�̃n���h����ǂݍ��܂���ƁA����XLIM��ǂݎ����
% �@�`�悳��Ă���͈͓��̃f�[�^��Excel�ɓ\��t����\�t�g�B
% �@�o�͂�Figure.xlsx�@���邢�́@{FigureName}.xlsx 
%   ���n��f�[�^����2DData�̂ݗL��
% 
%   Input:
%               fig     --- figureHandler 
%               Name    --- xlsxname ex. 'test' or 'test.xlsx'
%               
% 

%check filename
if nargin == 2
    filename = varargin{1};
    if ~strcmp(filename(end-4:end),'.xlsx')
        filename = [filename,'.xlsx'];
    end
elseif nargin == 1
    if isempty(fig.Name)
        filename = 'Figure.xlsx';
    else
        filename =[fig.Name,'.xlsx'];
    end
else
    disp('Error in FileName')
    return;
end

%Make CutData
FigureNumber =  size(fig.Children,1); %Figure��̃O���t�̐�

%initialize
legendFlag = zeros(FigureNumber,1);
XYDataCount = 1;

for i = 1:FigureNumber
    %judge legend    
    YaxisDataNumber =  size(fig.Children(i).Children,1);

    if YaxisDataNumber > 0 %Has Y-axis equal Data
        OutputData{XYDataCount} = ExtractXYData(fig,i);
        XYDataCount = XYDataCount + 1;
    elseif YaxisDataNumber == 0 %Legend
        OutputLegend{XYDataCount} = fig.Children(i).String;
        legendFlag(XYDataCount) = 1;
    else
        disp('Errro Occur@exportFigureData2Excel');
        disp('Check fig.Children');
    end
end

%Export to Excel
CellOffset = 0; %initialize
for k  = XYDataCount-1:-1:1
    %saveData
    SaveData = OutputData{k}; %adapt reverse numbering
    %Paste Cell
    xlRange = [char('A'+CellOffset) '2'];
    xlswrite(filename,SaveData,'Sheet1',xlRange)

    %Save and Paste Legend
    if legendFlag(k)
        SaveLegend =  [' ',OutputLegend{k}]; %Combine XData
        xlRange = [char('A'+CellOffset) '1'];
        xlswrite(filename,SaveLegend,'Sheet1',xlRange)
    end

    %Offset Update
    CellOffset = CellOffset + size(SaveData,2);
end

end

%Private Function
function CutoutXYData = ExtractXYData(fig,figNumber)
% figureNumber�̐}��XY�f�[�^��XLIM�Ő؂��茋�����ĕԂ��֐�
%initialize
OutputXData = [];
OutputYData = [];

YaxisDataNumber =  size(fig.Children(figNumber).Children,1);

for j = 1:YaxisDataNumber
    OutputYData(:,j) = fig.Children(figNumber).Children(j).YData;
end
OutputXData(:,1) =  fig.Children(figNumber).Children(j).XData;

XlimData = fig.Children(figNumber).XLim;

%Serch FistPoint & EndPoint
StartPoint = find(OutputXData >= XlimData(1),1,'first');
EndPoint   = find(OutputXData <= XlimData(2),1,'last');

%Cut and Combine Data
CutoutXYData = [OutputXData(StartPoint:EndPoint,1),...
                OutputYData(StartPoint:EndPoint,:)];
end

