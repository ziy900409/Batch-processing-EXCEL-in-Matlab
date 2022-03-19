%% Automatic extract column from excel files
% Frist, select the path of input files folder
% Second, input which columns do you need
% Third, select the path of output files folder
% then, the program will auto extract data from excel and deal with all of excel files in the folder
% Cloums input can using 1, 3, 5 to represent cloumns 
% or using B2:B4 to represent excel cloumns
% Important!!!
% Check variable "Files" in Workspace, because the order of files may inconsistent with your origin folder  
% Alternative plan, comment line 48 and Umcomment line 50, just add an append behind the origin files name,
% example: YourFileName_ed
% Author: Hsin Yang, 20210225
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

x = str2num(answer{1});%convert string in the answer to number
t = isempty(x(1));%judge the answer from dialog is number or not

if t == 0
    
    LengthAnswer = length(x);
else
    StringName = split(answer{1});%split string
    LengthString = length(StringName);
end
%% setting file path
path = strcat(input_folder, '\');
% return all of xlsx files of the folder
Files = dir(strcat(path,'*.xlsx'));
LengthFiles = length(Files);

%% extract cloumn from excel files

if t == 0
    for i = 1:LengthFiles
        [number1, text1, rawData1] = xlsread(strcat(path,Files(i).name));
        y = {};%setting cell matrix
    for j = 1:LengthAnswer
        y(:, j) = rawData1(:, x(j));
    end
    % setting output files name
        str = answer{2};
        OutputFileName = sprintf('%s_%g', str, i);
    %creat a new files name
    %stract: Concatenate strings horizontally
    % xlswrite(strcat(output_folder, '\', OutputFileName, '.xlsx'), y);
    %using input files name and append 'ed' behind the origin files name
        xlswrite(strcat(output_folder, '\', Files(i).name, '_ed', '.xlsx'), y);
    end
else
     for i = 1:LengthFiles
        yy = {};%setting cell matrix
        for k = 1:LengthString
        [number2, text2, rawData2] = xlsread(strcat(path,Files(i).name), StringName{k});
        yy(:, k) = rawData2;
        end
    % setting output files name
        str = answer{2};
        OutputFileName = sprintf('%s_%g', str, i);
    %creat a new files name
    %stract: Concatenate strings horizontally
    % xlswrite(strcat(output_folder, '\', OutputFileName, '.xlsx'), y);
    %using input files name and append 'ed' behind the origin files name
        xlswrite(strcat(output_folder, '\', Files(i).name, '_ed', '.xlsx'), yy);
     end
end
