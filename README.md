# Extraction.OCR

## 介绍
使用 Onenote API 实现的 图片OCR识别类

## 实现
- Onenote 2007 使用MODI组件进行OCR识别
- Onenote 2013 之后的版本使用指定XML文件进行OCR识别: 通过XML将图片写入Onenote的Page页，然后通过XML读取识别后的文本

## 注意
1. office2007需要安装office sp2补丁
2. 关闭onenote.dll的 嵌入互操作类型
3. 如果是在服务器上使用:
    - 开启 桌面体验 功能
    - 如果在服务器上使用onenote 2007: 由于该组件是32位的，所以调用该接口的程序必须也是32位程序
    - 如果是在服务器上使用onenote 2010，出现了 `Retrieving the COM class factory for component with CLSID {D7FAC39E-7FF1-49AA-98CF-A1DDD316337E} failed due to the following error: 80080005 Server execution failed (Exception from HRESULT: 0x80080005 (CO_E_SERVER_EXEC_FAILURE))` 这个异常，请在调用此dll的服务登陆身份修改为管理员账号，并且用此账号创建onenote笔记本，另外，本示例代码需要笔记本格式为 onenote2007

## ETC
- 开源的OCR库: 可以试下 Tesseract ，一个开源的，由谷歌维护的OCR软件
- [错误代码 (OneNote 2013)](https://msdn.microsoft.com/zh-cn/library/jj680117)
