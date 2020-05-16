1.python下载地址：https://www.python.org/downloads/release/python-383/
2.安装openpyxl 
    ctrl+R cmd 
        输入:pip install openpyxl
3.打开setExecl.py 修改路径
    path为项目路径

    wb = load_workbook(path+'xxx')为 危险源辨识与风险评价表
    wb2 = load_workbook(path+'2/demo.xlsx')为 作业活动分级管控表

    生成路径在最底部 为final.xlsx
4.使用
    setExecl.py右键打开方式=>Python