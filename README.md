# ContractEntry


功能：
1，word录入web：将本地Word文件，录入web页面富文本框
2，Word对比：对两个Word文件进行对比，支持章节，分段对比
3，支持3个模式选择：A：Word录入+Word文件比对   B：只对两个Word文件进行比对   C：只对Word进行web录入
4，支持多文件录入和对比
5，支持替换Word文件里的数据，替换规则是写死的，更改需要联系脚本开发者
6，录入的结果展示在result文件夹里的result.log文件：只记录了和下载文件的对比结果


使用方法：
1，在Code文件夹里的 config.py文件中，填写参数项
2，在项目根目录cmd运行【python run.py】,脚本开始工作


底层：
1，python
2，selenium
3，python-docx


