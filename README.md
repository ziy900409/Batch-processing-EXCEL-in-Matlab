# Batch-processing-EXCEL-in-Matlab
Frist, select the path of input files folder
Second, input which columns do you need
Third, select the path of output files folder
then, the program will auto extract data from excel and deal with all of excel files in the folder

Cloums input can using 1, 3, 5 to represent cloumns, number between number using ',' to separate its

OR using B2:B4 to represent EXCEL cloumns, string between string using 'space' to separate its

Important!!!
Check variable "Files" in Workspace, because the order of files may inconsistent with your origin folder  
Alternative plan, comment line 35 and Umcomment line 37, just add an append behind the origin files name, example: YourFileName_ed
If you want different appendent, just change '_ed' in line xlswrite.

批次處理EXCEL,使用提取資料為範例
1. 選擇input folder的路徑
2. 選擇output folder的路徑
3. 輸入想提取的欄位編號以及預計output file的名稱
程式會自動將output file名稱後續編碼
例如：OutputFileName_1, OutputFileName_2, OutputFileName_3...
需注意Workspace裡面Files的變數名稱，由於matlab讀取當按順序的問題，可能會將10擺在2的前面
例如：1>10>11>2>3>4>5

解決方法：直接使用原本的檔案名稱，並在結尾加_ed，可避免檔案搞混

Comment line xlswrite(strcat(output_folder, '\', OutputFileName, '.xlsx'), y);

Uncomment line xlswrite(strcat(output_folder, '\', Files(i), '_ed', '.xlsx'), y);

--------------------Update---------------------------------------------------------------------------------------------------
Code_2 2021.03.03
新增一種輸入欄位方式，現在欄位可用數字以及excel預設的欄位來輸入
1. 輸入欄位編號，不同欄位使用逗號(comma)分開，例如：1, 2, 3....
2. 輸入excel欄位與範圍，不同欄位使用空格(space)分開，例如：A1:A500 B1:B500.....，整欄資訊 C:C D:D E:E....

捨棄output file名稱，會讓使用者不好尋找原始檔案

Code_3 2021.03.04
新增批次處理多層檔案夾的方法
FolderList: 回傳所有該檔案夾下所有檔案名稱及路徑
再創建新檔案夾，遍歷所有檔案
