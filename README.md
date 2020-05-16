1.python下载地址：https://www.python.org/downloads/release/python-383/
2.安装openpyxl 
    ctrl+R cmd 
        输入:pip install openpyxl
3.另存文件格式
	文件格式必须是xlsx，不可改后缀，必须另存
4.打开setExecl.py 修改路径
# 项目路径
    1.危险辨识表(最好为英文路径)
        wxPath = "C:/Users/79234/Desktop/py/1/demo.xlsx"
    2.风险管控表(最好为英文路径)
        fxPath = "C:/Users/79234/Desktop/py/2/demo.xlsx"
    3.最终生成的文件目录(最好为英文路径)
        finalPath = "C:/Users/79234/Desktop/py/final.xlsx"
# 标头行数修改
	代码62行     默认删除前4行表头
        worksheet.delete_rows(1, 4)
        worksheet2.delete_rows(1, 4) 
5.使用
    打开最终的excel时，不可执行文件
    setExecl.py右键打开方式=>Python