# Word2MD 📝➡️ Markdown

一个简单好用的 Word 文档（.docx/.doc）转 Markdown 工具，支持图片、公式（LaTeX）、表格与批量转换。提供图形界面与一键可执行程序，开箱即用。

## ✨ 特性
- ✅ 支持 .docx/.doc（若无 Pandoc 则自动回退方案）
- 🧮 保留数学公式（LaTeX 语法，原样输出）
- 🖼️ 自动提取并保存图片，Markdown 中正确引用
- 📊 表格结构保留，转换为 Markdown 表格
- 📂 支持单文件与批量转换
- 🪟 现代化 GUI（进度条、状态提示、日志）
- 🚀 提供 Windows 独立可执行文件，无需安装 Python/Pandoc

## 📦 目录结构
```
github_release/
├─ README.md                 # 项目说明（本文件）
├─ word2md_enhanced.py       # 稳定版源代码（GUI入口）
├─ requirements.txt          # 依赖列表（用于源码运行）
├─ Word2MD.spec              # PyInstaller 打包配置
├─ build_exe.py              # 打包脚本（Python 版）
├─ build.bat                 # 打包脚本（Windows 快速版）
├─ word2md_icon.ico          # 应用图标（用于打包/展示）
├─ word2md_icon.png          # 应用图标PNG
├─ bin/
│  └─ Word2MD.exe           # 可直接运行的 Windows 可执行文件
└─ examples/
   └─ X射线脉冲星光子到达时间建模.docx   # 示例文档
```

## 🖱️ 使用方式

- 方式一：直接运行可执行文件（推荐）
  1. 双击 `bin/Word2MD.exe`
  2. 在界面中选择转换模式（单文件/批量）
  3. 选择输入文件与输出目录
  4. 点击「开始转换」，等待完成 ✅

- 方式二：源码运行（需要 Python 3.9+）
  ```bash
  pip install -r requirements.txt
  python word2md_enhanced.py
  ```

## 🧠 设计说明
- 默认优先使用 Pandoc 进行高质量转换；若系统未安装 Pandoc，会自动回退至 Mammoth + 自定义处理，确保可用性。
- 公式维持原有 LaTeX 语法，不进行二次改写，避免渲染差异。
- 图片提取到输出目录的 `images/` 子目录，并在 Markdown 中以相对路径引用。
- 批量模式下显示每个文件的处理状态与进度。

## 🔨 自行打包（可选）
你可以使用以下任一种方式打包 exe：

- 使用批处理脚本：
  ```bat
  build.bat
  ```

- 使用 Python 脚本：
  ```bash
  python build_exe.py
  ```

构建产物默认位于 `dist/Word2MD.exe`，本仓库已提供一份在 `bin/` 下的现成可执行文件。

## 📚 示例与测试
- 示例文档位于 `examples/` 目录。
- 你可以将自己的 `.docx` 文件用于测试，转换后的 Markdown 和图片将输出到你指定的目录。

## ❓常见问题（FAQ）
- Q: 一定需要安装 Pandoc 吗？
  - A: 不需要。若未安装，程序会自动使用备用方案，功能仍可用。
- Q: 公式为什么不被修改？
  - A: 为保持渲染一致性，本工具按原有 LaTeX 语法直接输出。
- Q: 转换很慢怎么办？
  - A: 大文档或含大量图片/表格时较慢，请耐心等待；界面会显示进度。

## 📄 许可证
建议使用 MIT License（可根据你的需求调整）。

```
MIT License

Copyright (c) 2024

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 🌟 致谢
- 感谢开源社区的优秀项目：`python-docx`、`mammoth`、`pypandoc`、`Pillow`、`lxml` 等。

---

如果本项目对你有帮助，欢迎 ⭐️ Star 支持！
