# Batch-processing-EXCEL-in-Matlab
Frist, select the path of input files folder
Second, input which columns do you need
Third, select the path of output files folder
then, the program will auto extract data from excel and deal with all of excel files in the folder


Important!!!
Check variable "Files" in Workspace, because the order of files may inconsistent with your origin folder  
Alternative plan, comment line 35 and Umcomment line 37, just add an append behind the origin files name, example: YourFileName_ed

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
