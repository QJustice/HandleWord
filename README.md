# 批量替换word文本框文字

## 使用前做准备

样例word如下：

![样例word](https://github.com/qascetic/HandleWord/blob/master/READMEFiles/word_image.jpg "word")

运行程序可获得以下界面：

![运行图](https://github.com/qascetic/HandleWord/blob/master/READMEFiles/run.jpg "初次运行")

## 使用方法

1. 在需要替换的word文本框内容位置上写上标记符号，样例中标记符号NAME，然后保存好文件。**完成操作后必须关闭word，否则程序没有写入word的权限**

2. 运行程序
   
   + 姓名标志栏写入刚才在word中写的标记符号
   
   + 输入模板word路经或点击右则三个点选择目标模板word，程序会自动填写模板word路径。
   
   + 确定好分隔符，在程序分隔符一栏写入，样例中用的是|，然后输入获奖者姓名，若有多个请用分隔符分割。例如，小明|小王|张三。程序能自动识别获奖者姓名并会把获奖人数显示在获奖人数栏，供用户参考。
   
   + 决定好生成的文件夹名并写入到文件夹名栏。随后生成的word会保存在此文件夹。
   
   + 最后点击”生成“按钮系统会询问是否开始并提示识别到的获奖人名数量供用户参考。若点击是即可完成批量生成，若点击否系统会取消生成操作。

3. 其他功能介绍
   
   + “清空”按钮：点击即可清空获奖名单输入框
   
   + “关闭”按钮：点击即可关闭窗口
   
   ![运行效果图](https://github.com/qascetic/HandleWord/blob/master/READMEFiles/run_image.jpg)

## 注意事项

本程序只能修改word中文本框内的文字。无法修改段落内的文字。需要批量生成的word需要进行相关适配。
