批量把word文件转换为pdf

解压w2p.zip, 把二进制文件放在需要转换的文件夹中

执行w2p.exe，会把当前文件夹中所有doc docx文件转换为pdf文件

ps：同名pdf文件会被删除，请注意


1.1:
--------
fix bugs：
1. 修复了因文件名中出现"."导致的文件名截取不完整
2. 排除了当word文件打开时出现的缓存文件也被转码导致出错
3. 转码完成后，会关闭word进程

## License

本项目采用 [![license](https://img.shields.io/github/license/littlemo/mohand.svg)](https://github.com/littlemo/mohand) 协议开源发布，请您在修改后维持开源发布。
