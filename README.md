## ToolHost

* 工具容器，可插拔工具箱子

## ToolCore 

* 工具接口及工具管理核心方法

##  MODBUSCONFIG

* elc 卡件 modbusTcp配置批量生成
##  tempBatch 

* 根据模板批量生成文件 参数占位符 #{任意字符} 例如：#{P}

##  StatsBatch 

* stasts 文件转EXCEL

## Plugin格式
* Plugin的csproj格式如下：

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <UseWPF>true</UseWPF>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\..\Toolbox.Core.csproj" />
  </ItemGroup>

</Project>
```
* 更改格式后重新编译，并把dll文件复制到plugins文件夹即可

* 部署后的Plugins文件夹结构
```bash
..\2317\ToolboxHost\bin\Debug\Plugins\TestPlugin1\
├── TestPlugin1.dll         # 编译后的插件程序集
├── plugin.json            # 插件配置文件
├── Dependencies\          # 依赖文件夹
└── Resources\             # 资源文件夹
```
