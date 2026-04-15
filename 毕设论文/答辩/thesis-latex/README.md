# 上海理工大学学位论文 LaTeX 模板

## 使用方法（Overleaf）

1. 将 `main.tex` 和 `references.bib` 打包为 `.zip` 上传至 [Overleaf](https://www.overleaf.com)
2. 在 Overleaf 的 **Menu → Compiler** 中选择 **XeLaTeX**
3. 点击 **Recompile** 即可编译

## 文件说明

| 文件 | 说明 |
|---|---|
| `main.tex` | 论文主文件，包含完整结构 |
| `references.bib` | 参考文献数据库（可选，也可用内联 thebibliography） |

## 格式规范依据

- 《上海理工大学博士、硕士研究生学位论文写作规范（2018年修订版）》
- GB/T 7713—1987《科学技术报告、学位论文和学术论文的编写格式》
- GB/T 7714—2015《信息与文献 参考文献著录规则》

## 关键格式速查

| 项目 | 规格 |
|---|---|
| 纸张 | A4 |
| 页边距 | 上3.5cm 下2.5cm 左3.0cm 右3.0cm |
| 页眉 | 五号宋体居中，奇偶页不同 |
| 页码 | 五号 Times New Roman 居中 |
| 章标题 | 黑体小二号居中，段前24pt段后18pt |
| 一级节标题 | 宋体四号加粗，段前24pt段后6pt |
| 二级节标题 | 宋体小四号加粗，段前12pt段后6pt |
| 正文 | 宋体小四号，首行缩进2字符，行距固定值20pt |
| 图名 | 宋体小四号居中，图下方段前6pt段后12pt |
| 表名 | 宋体小四号居中，表上方段前12pt段后6pt |
| 参考文献 | 宋体小四号，行距固定值16pt |

## 字体说明

模板默认使用 Fandol 免费字体（Overleaf 自带），无需额外安装。

如果在本地编译并需要更丰富的字体，可修改 `main.tex` 中的字体设置：

```latex
% 本地编译可使用系统字体
\setCJKmainfont{SimSun}       % 宋体
\setCJKsansfont{SimHei}       % 黑体
\setCJKmonofont{FangSong}     % 仿宋
\setmainfont{Times New Roman} % 英文正文
```

## 常用操作

### 插入图片
```latex
\begin{figure}[htbp]
  \centering
  \includegraphics[width=0.8\textwidth]{figures/your_image.png}
  \caption{图片标题}
  \label{fig:your_label}
\end{figure}
```

### 插入三线表
```latex
\begin{table}[htbp]
  \centering
  \caption{表格标题}
  \label{tab:your_label}
  \zihao{5}\songti
  \begin{tabular}{ccc}
    \toprule
    列1 & 列2 & 列3 \\
    \midrule
    数据 & 数据 & 数据 \\
    \bottomrule
  \end{tabular}
\end{table}
```

### 插入数学公式
```latex
% 行内公式：$E = mc^2$
% 居中公式：
\begin{equation}
  f(x) = \int_{-\infty}^{\infty} \hat{f}(\xi) e^{2\pi i \xi x} d\xi
  \label{eq:fourier}
\end{equation}
```

### 添加新章节
```latex
\chapter{章节标题}
\section{一级节标题}
\subsection{二级节标题}
正文内容...
```

## 注意事项

1. 封面由学院统一发放，本模板提供的是内封面信息模板
2. 版权使用授权书需从研究生院网站下载替换
3. 目录由 LaTeX 自动生成，无需手动编写
4. 如使用 `references.bib`，取消 main.tex 中 `\bibliography{references}` 的注释即可
5. 编译方式必须选择 **XeLaTeX**（因为涉及中文字体）
