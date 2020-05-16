1.python下载地址：https://www.python.org/downloads/release/python-383/
2.安装openpyxl 
    ctrl+R cmd 
        输入:pip install openpyxl
3.另存文件格式
	文件格式必须是xlsx，不可改后缀，必须另存
4.打开setExecl.py 修改参数
# 项目路径修改
	1.风险表(最好改为英文)
		riskPath = "C:/Users/79234/Desktop/examine/risk.xlsx"
	2.排查表(最好改为英文)
		pcPath = "C:/Users/79234/Desktop/examine/data.xlsx"
	3.最终生成的目录(最好改为英文)
		finalPath = "C:/Users/79234/Desktop/examine/wb.xlsx"
# 标头行数修改
	代码70行 exWS.delete_rows(1, 3) 默认3行, 3改为x行
	表格尾部
		代码68行	exWS.delete_rows(exMaxCol-4, 5) 默认最后五行没用
	确定表格数据
		风险点 B列	
		项目 C列	
		内容 D列	
		"导致事故的原因(危害因素)" E列
		排查内容 G列
# 排查单位 单元格必须在D2或（合并单元格情况下）占有D2

5.使用
	打开最终的excel时，不可执行文件
	examine.py右键打开方式=>Python
