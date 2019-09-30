# 代码仓库
构建自己的代码仓库

### python对excel的操作

- 背景：

```
有两个excel文件：f1.xlsx和f2.xlsx
```

| f1.xlsx文件为：                                   | f2.xlsx文件为：                                   | 操作之后，f2.xlsx文件内容变为：                   |
| ------------------------------------------------- | ------------------------------------------------- | ------------------------------------------------- |
| ![1569852766602](README.assets/1569852766602.png) | ![1569852799960](README.assets/1569852799960.png) | ![1569853016484](README.assets/1569853016484.png) |

- 
  要点：

```
1. 实现python对xlsx的读写操作，数据量18万
2. 用到的库有：xlrd,xlutils,openpyxl,datetime
3. python格式化日期时间输出
```

- 实现：

python实现：[code](./python_excel/rw_excel.py) 

```
1. 函数compared_data(readfile, writefile)：
      功能：完成对文件f1和f2的读取工作,并将f1的时间添加到f2中，保存到新的文件f2_mid.xlsx中
      输入：readfile:f1; writefile:f2
      输出：f2_mid
2. 函数add_time(filename):
	  功能：将f2中的空白时间补充完整
	  输入：f2_mid.xlsx
	  输出：new_f2.xlsx
```
​		

