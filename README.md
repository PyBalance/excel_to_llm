# Excel Analyzer for LLM Understanding

## English

Excel Analyzer is a Rust-based desktop application designed to process Excel files and generate structured output that can be easily consumed by Large Language Models (LLMs). This tool helps bridge the gap between raw Excel data and LLM input, facilitating better understanding and analysis of spreadsheet content.

Key features:
- Load and analyze multiple Excel files
- Extract sheet names, headers, and sample data
- Customizable sample size (number of rows to display)
- Multiple output formats (Markdown, XML, Plain Text) suitable for LLM ingestion
- Multi-threaded analysis for improved performance
- User-friendly GUI built with egui

Usage:
1. Run the application
2. Add Excel files for analysis
3. Set the number of sample rows
4. Choose an output format
5. Click "Analyze Files"

The structured output can then be used as input for LLMs, enabling them to better understand and work with Excel file contents.

## 中文

Excel 分析器是一个基于 Rust 的桌面应用程序，旨在处理 Excel 文件并生成结构化输出，以便大型语言模型（LLM）可以轻松使用。该工具有助于缩小原始 Excel 数据和 LLM 输入之间的差距，促进对电子表格内容的更好理解和分析。

主要功能：
- 加载和分析多个 Excel 文件
- 提取工作表名称、表头和样本数据
- 可自定义样本大小（显示的行数）
- 多种适合 LLM 摄取的输出格式（Markdown、XML、纯文本）
- 多线程分析以提高性能
- 使用 egui 构建的用户友好界面

使用方法：
1. 运行应用程序
2. 添加要分析的 Excel 文件
3. 设置样本行数
4. 选择输出格式
5. 点击"Analyze Files"

生成的结构化输出可以作为 LLM 的输入，使其能够更好地理解和处理 Excel 文件内容。

## Installation

1. Ensure you have Rust installed on your system.
2. Clone this repository.
3. Run `cargo build --release` in the project directory.
4. The executable will be available in `target/release/`.

## Dependencies

- egui: GUI framework
- calamine: Excel file reading
- rfd: File dialog
- strum: Enum utilities

For a complete list of dependencies, please refer to the `Cargo.toml` file.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the [MIT License](LICENSE).