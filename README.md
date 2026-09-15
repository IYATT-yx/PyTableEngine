# PyTableEngine 思能快表引擎

## 背景 / Background

2026/9/3  
PyTableEngine 计划开发作为一个集中的 Excel 插件管理平台，插件采用 Python 开发，可以通过源码或二进制包的形式进行分发。  

我其实很希望 Excel 可以官方支持本地执行 Python 脚本，可惜没有。我去年尝试过 xlwings，我折腾了半天多都没搞成功，最后还是放弃了。
前段时间我也通过 AI 帮我写 VBA 实现了一些便捷功能，但 VBA 主要问题是和文件关联紧密，过段时间再想用就找不到了，而且编写调试也不方便。xlwings 的工作方案和 VBA 相似，只是支持了 Python，但是也是一样的毛病，要嵌入文件中。再加上当时我折腾 xlwings 很久都没成功，还需要自己去安装 Python，关键还没搞成功，整体看易用性并不高。  

我上周拟定计划准备开发这样一个工具：一个中心引擎，这个引擎自身不提供实际功能，而是作为一个桥梁和 Excel 建立连接，具体功能实现通过 Python 开发插件实现，并且这个引擎是通过二进制可执行文件发布，不需要用户也安装 Python，也能直接运行插件，简化使用。这样具体的功能就和 Excel 文件解耦了，可以随时方便使用。我拖到今天才开始动工，并新增了一个设想：VBA 是可以自定义公式的，可以在单元格中调用。按前期设想的框架就做不到这样的功能，但是我并不想直接以公式的方式嵌入单元格，这样文件一旦移动到没有本引擎的电脑上就就变成`#NAME?`。我希望单元格是显示静态结果的，Excel 文件又支持嵌入自定义数据，那么可以考虑在 Excel 中嵌入不可见的计算图，指导某插件提供的“公式”使用哪些单元格，结果写入到哪些单元格，在引擎工作时，也能做到单元格自动调用“公式”计算更新。  

---

2026/9/3  
The PyTableEngine project was planned as a centralized Excel add-in management platform. The add-ins are developed in Python and can be distributed either as source code or binary packages.

I really wished Excel officially supported running local Python scripts, but unfortunately, it doesn't. I tried xlwings last year, spent over half a day setting it up without success, and eventually gave up.
Recently, I used AI to help me write VBA to implement some handy features. However, the main problem with VBA is that it is tightly coupled with individual files—if I want to use it again after a while, I can't find it, not to mention that writing and debugging it are inconvenient. The workflow of xlwings is similar to VBA; although it supports Python, it suffers from the same flaw of requiring embedding within files. Combined with the fact that I spent a long time trying to set up xlwings without success, and that it required manually installing Python, its overall usability is quite low.

Last week, I outlined a plan to develop a tool: a central engine that provides no actual functionality itself, but serves as a bridge to establish a connection with Excel. Concrete functions are implemented through Python-developed add-ins. Moreover, this engine is distributed as a binary executable file, eliminating the need for users to install Python while allowing them to run add-ins directly to simplify usage. This decouples specific functionalities from Excel files, making them convenient to use at any time. I put off starting the development until today and came up with an additional idea: VBA allows defining custom formulas that can be called within cells. The initial framework could not support such a feature. However, I don't want to embed formulas directly into cells, because if the file is moved to a computer without this engine, it will result in a `#NAME?` error. I want the cells to display static results. Since Excel files support embedded custom data, I can consider embedding an invisible computation graph in Excel. This graph directs the "formulas" provided by a specific add-in regarding which input cells to read and which output cells to write to. When the engine is running, it can also automatically trigger these "formulas" to compute and update cell values.

## 测试环境 / Testing Environment

* Python 3.14.5  （Windows 10 及以上可用）  
* Microsoft Office 专业增强版 2024  
* Visual Studio Community 2026（Nuitka 打包时使用）

---

* Python 3.14.5 (Available on Windows 10 and above)
* Microsoft Office Professional Plus 2024
* Visual Studio Community 2026 (Used during Nuitka packaging)

## 插件 / Plugins

### 插件基本规则 / Basic Rules for Plugins

插件必须被包含在一个文件夹下，文件夹命名为：作者名_插件名，该名称称为插件 ID，保证插件的唯一性（区分大小写）。插件的入口文件命名也是 ID，且入口文件中必须包含匹配的元数据，匹配作者名和插件名。插件文件夹需放置在 `extensions` 目录下。  
例如我的插件名为 `Demo`，作者名为 `IYATT`，那么必须有以下结构：  
```text
extensions
 |
 |——IYATT_Demo
        |——IYATT_Demo.py/.pyd
```

且 IYATT_Demo.py/.pyd 中必须存在全局变量 `pluginInfo`，格式如下：
```python
pluginInfo = {
    'name': 'Demo',
    'author': 'IYATT',
    'description': '<插件的描述信息>',
    'version': '<插件的版本>',
}
```

可以手动分发插件复制到 `extensions` 目录下，启动工具即可识别。  

---

Plugins must be contained within a folder, named: author_name_plugin_name, which is called the plugin ID, ensuring the uniqueness of the plugin (case-sensitive). The entry file of the plugin must also match the author name and plugin name. The plugin folder must be placed in the `extensions` directory.  
For example, if my plugin name is `Demo` and my author name is `IYATT`, the structure must be as follows:  
```text
extensions
 |
 |——IYATT_Demo
        |——IYATT_Demo.py/.pyd
```

And the IYATT_Demo.py/.pyd file must contain a global variable `pluginInfo` with the following format:  
```python
pluginInfo = {
    'name': 'Demo',
    'author': 'IYATT',
    'description': '<Description of the plugin>',
    'version': '<Version of the plugin>',
}
```

Plugins can be manually distributed by copying them to the `extensions` directory and starting the tool to recognize them.  

### 在线插件 / Online Plugins

软件首次启动会创建`config.ini`，在 `[market]` 下 `repo_index =` 后可填写索引地址，默认不填写时为 GitHub 插件中心仓库索引。  
GitHub 插件中心仓库索引：  
```text
https://raw.githubusercontent.com/IYATT-yx/PyTableEnginePluginRepository/main/index.json
```

Gitee 插件中心仓库仓库索引：  
```text
https://gitee.com/iyatt/PyTableEnginePluginRepository/raw/main/index.json
```

发布插件只能前往[GitHub 插件中心仓库](https://github.com/IYATT-yx/PyTableEnginePluginRepository)，[Gitee 插件中心仓库](https://gitee.com/iyatt/PyTableEnginePluginRepository) 仅同步，不支持提交插件注册。  

---

The software creates `config.ini` upon its first startup. The index address can be filled in under `[market]` with `repo_index =` after the default value is not filled in.  
GitHub Plugin Center Repository Index:  
```text
https://raw.githubusercontent.com/IYATT-yx/PyTableEnginePluginRepository/main/index.json
```

Gitee Plugin Center Repository Index:  
```text
https://gitee.com/iyatt/PyTableEnginePluginRepository/raw/main/index.json
```
Plugins can only be published to the [GitHub Plugin Center Repository](https://github.com/IYATT-yx/PyTableEnginePluginRepository). The [Gitee Plugin Center Repository](https://gitee.com/iyatt/PyTableEnginePluginRepository) only synchronizes and does not support submitting plugin registration.  

## 许可协议 / License

本项目基于 [MIT License](LICENSE) 开源。

---

This project is open source under the [MIT License](LICENSE).