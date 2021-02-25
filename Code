%% create dialog
% setting path of input and output floder
input_folder = uigetdir('C:\', 'Select the Path of input floder');
output_folder = uigetdir('C:\', 'Select the Path of output floder');
% input which cloumns do you want
prompt = {'Cloumn (Please use ","  to seperate different number)','name of output file'};
dlgtittle = 'input which columns do you want (Integer only!!)';
dims = [1 100];
definput = {'5, 6, 7', 'OutFileName'};
answer=inputdlg(prompt, dlgtittle, dims,definput);
%convert string to number
x = str2num(answer{1});


%% setting file path
path = strcat(input_folder, '\');
% return all of xlsx files of the folder
Files = dir(strcat(path,'*.xlsx'));
LengthFiles = length(Files);
LengthAnswer = length(x);

%% extract cloumn from excel files
for i = 1:LengthFiles
[number, text, rawData] = xlsread(strcat(path,Files(i).name));
y = [];
for j = 1:LengthAnswer
    z = number(:, x(j));
    y(:, j) = z;
end
% setting output files name
str = answer{2};
OutputFileName = sprintf('%s_%g', str, i);
%creat a new files name
%stract: Concatenate strings horizontally
xlswrite(strcat(output_folder, '\', OutputFileName, '.xlsx'), y);
%using input files name and append 'ed' behind the origin files name
%xlswrite(strcat(output_folder, '\', Files(i), '_ed', '.xlsx'), y);
end
