# 台湾地质钻孔数据批量下载工具

## 安装步骤

1. 确保已安装Node.js（v16或更高版本）
2. 解压下载的zip文件
3. 在命令行中进入解压后的目录
4. 运行以下命令安装依赖： npm install
5. 安装Playwright浏览器：npx playwright install chromium
## 使用方法

1. 运行爬虫：npm start
2. 从已有日志导入已完成项目ID：npm start -- --import-log 日志文件路径
## 配置文件

配置文件位于`drill_config`目录，数据保存在`drill_data`目录。
