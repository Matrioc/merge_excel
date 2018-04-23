#### 把多个excel文件的多个sheet表单的内容合并

- merge_excel.py的功能是简单的按sheet表单的数字索引合并，即按excel的第几个sheet内容合并，所以如果对于excel里对应数字缩影的sheet名字不同则合并时的内容会造成错误。

- 因此merge_excel_shname.py的作用则是解决该问题，是按sheet的名字进行内容合并，即对于多个excel里具有相同sheet-name的内容进行合并。
