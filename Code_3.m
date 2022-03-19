%% Automatic extract column from excel files
% Frist, select the path of input files folder
% Second, input which columns do you need
% Third, select the path of output files folder
% then, the program will auto extract data from excel and deal with all of excel files in the folder
%
% Cloums input can using 1, 3, 5 to represent cloumns, number between
% number using ',' to separate its
% OR using B2:B4 to represent EXCEL cloumns, string between string using
% 'space' to separate its
% Important!!!
% Check variable "Files" in Workspace, because the order of files may inconsistent with your origin folder  
% Alternative plan, comment line 72 and Umcomment line 74, just add an append behind the origin files name,
% example: YourFileName_ed
%
% Author: Hsin Yang, 20210304
clc, clear

%% create dialog
% setting path of input and output floder
input_folder = uigetdir('D:\', 'Select the Path of input floder');
output_folder = uigetdir('D:\', 'Select the Path of output floder');
% input which cloumns do you want
prompt = {'Cloumn (Please use ","  to seperate different number)','name of output file'};
dlgtittle = 'input which columns do you want (Integer only!!)';
dims = [1 100];
definput = {'B2:B10 C2:C10', 'OutFileName'};
answer=inputdlg(prompt, dlgtittle, dims,definput);

%% judge information of input columns is string or number
x = str2num(answer{1});%convert string in the answer to number
t = isempty(x(1));%judge the answer from dialog is number or not

if t == 0
    LengthAnswer = length(x);
else
    StringName = split(answer{1});%split string
    LengthString = length(StringName);
end

%% setting path of file and folder 
FolderList = dir(input_folder);% return all of folder list
%remove '.', '..' in the FolderList
FolderList = FolderList(~ismember({FolderList.name}, {'.', '..'}));
LengthFolderList = length(FolderList);

for k = 1:LengthFolderList
path2 = strcat(FolderList(k).folder, '\', FolderList(k).name, '\');

%% creat a new Folder, name follow origin folder and appendent '_ed'
ParentFolder = output_folder;
FolderName = FolderList(k).name;
status = mkdir(ParentFolder, FolderName); %creat a new folder

% return all of xlsx files of the folder
Files2 = dir(strcat(path2, '*.xlsx'));
LengthFiles = length(Files2);
%% extract cloumn from excel files

if t == 0
    for i = 1:LengthFiles
        [number1, text1, rawData1] = xlsread(strcat(path2,Files2(i).name));
        y = {};%setting cell matrix
    for j = 1:LengthAnswer
        y(:, j) = rawData1(:, x(j));
    end
    % setting output files name
        str = answer{2};
        OutputFileName = sprintf('%s_%g', str, i);
    %creat a new files name
    %stract: Concatenate strings horizontally
    % xlswrite(strcat(ParentFolder, '\', FolderName, '\', OutputFileName, '.xlsx'), y);
    %using input files name and append 'ed' behind the origin files name
        xlswrite(strcat(ParentFolder, '\', FolderName, '\', Files2(i).name, '_trim', '.xlsx'), y);
    end
else
     for i = 1:LengthFiles
        yy = {};%setting cell matrix
        for k = 1:LengthString
        [number2, text2, rawData2] = xlsread(strcat(path2,Files2(i).name), StringName{k});
        yy(:, k) = rawData2;
        end
    % setting output files name
        str = answer{2};
        OutputFileName = sprintf('%s_%g', str, i);
    %creat a new files name
    %stract: Concatenate strings horizontally
    % xlswrite(strcat(ParentFolder, '\', FolderName, '\' OutputFileName, '.xlsx'), y);
    %using input files name and append 'ed' behind the origin files name
        xlswrite(strcat(ParentFolder, '\', FolderName, '\', Files2(i).name, '_trim', '.xlsx'), yy);
     end
end

end
