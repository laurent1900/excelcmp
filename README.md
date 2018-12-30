# excelcmp
用途

对比两个Excel中的第一列，找出其中不同，输出为一个新的Excel。

1、环境要求

Python2

2、安装

pip install xlrd

pip install xlwt

参数解释：

'-o','--old', help="Enter the filename1 that you want to compare")

要进行比较的第一个文件

'-n','--new', help="Enter the filename2 that you want to compare")

要进行比较的第二个文件

'-s','--sheet', help="Enter the sheet number that you want to compare. exp: sheet1 -s 0")

要进行比较的sheet页，0代表第一个sheet页，1代表第二个sheet页，以此类推

'-c','--column', help="Enter the column number that you want to compare")

要进行比较的列数，0代表第一列，1代表第二列，一次类推
