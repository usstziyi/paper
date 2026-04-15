# 上海理工大学学位论文 LuaLaTeX 模板

## 使用方法（Overleaf）

1. 将 `main.tex` 和 `references.bib` 打包为 `.zip` 上传至 [Overleaf](https://www.overleaf.com)
2. 在 Overleaf 的 **Menu → Compiler** 中选择 **LuaLaTeX**
3. 点击 **Recompile** 编译

## 为什么选 LuaLaTeX 而非 XeLaTeX？

| 特性 | XeLaTeX | LuaLaTeX |
|---|---|---|
| CJK 支持 | xeCJK 宏包 | luatexja（原生 Lua 集成） |
| microtype 字间距优化 | 有限支持 | **完整支持**（expansion + protrusion） |
| 脚本扩展 | 不可用 | **Lua 脚本**可嵌入文档 |
| Unicode 处理 | 好 | **更好**（原生 UTF-8） |
| 编译速度 | 较快 | 略慢 |
| 字体管理 | fontspec（XeTeX 后端） | fontspec（LuaTeX 后端） |
| Overleaf 兼容 | ✅ | ✅ |
| 与 XeLaTeX 互换 | — | ✅（代码几乎相同） |

**简单说：LuaLaTeX 排版质量更高，XeLaTeX 兼容性更广。** 推荐优先用 LuaLaTeX。

## 文件说明

| 文件 | 说明 |
|---|---|
| `main.tex` | 论文主文件 |
| `references.bib` | 参考文献数据库（可选） |
| `README.md` | 本文件 |

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

### Overleaf（推荐方案）

模板默认使用 Noto CJK 系列 + TeX Gyre Termes（Times New Roman 兼容替代），Overleaf 自带，无需额外安装。

### 本地编译（更丰富的字体选择）

```latex
% 替换为系统字体
\setmainjfont{SimSun}[BoldFont=SimHei]     % 宋体 + 黑体
\setsansjfont{SimHei}                        % 黑体
\setmonojfont{FangSong}                      % 仿宋
\setmainfont{Times New Roman}                % 英文正文
```

## LuaLaTeX 独有功能

### microtype 字间距微调

```latex
% 已在模板中启用，无需额外操作
% LuaLaTeX 的 microtype 比 XeLaTeX 支持更完整
\UseMicrotypeSet[protrusion=basic, expansion=alltext]{lua}
```

### Lua 脚本嵌入

模板中已预留 `luacode` 环境，可编写 Lua 脚本：

```latex
\begin{luacode}
-- 示例：自定义章节统计
function count_chapters()
  -- Lua 代码
end
\end{luacode}
```

常用场景：
- 批量处理参考文献格式
- 动态生成图表编号
- 文件读取与数据处理
- 复杂的排版逻辑

## 三版本对比

| | Word 版 | XeLaTeX 版 | LuaLaTeX 版（本模板） |
|---|---|---|---|
| 编辑方式 | WYSIWYG | 代码 | 代码 |
| 学习曲线 | ⭐ | ⭐⭐⭐ | ⭐⭐⭐ |
| 排版质量 | 良好 | 优秀 | **最佳** |
| 数学公式 | 困难 | 极佳 | 极佳 |
| 参考文献管理 | 手动 | 自动 | 自动 |
| 目录自动生成 | 需手动更新 | ✅ | ✅ |
| 跨平台 | 需 Office | Overleaf | Overleaf |
| 推荐场景 | 导师要求 Word | 通用 | **追求最佳排版** |

## 注意事项

1. **编译器必须选 LuaLaTeX**，选 XeLaTeX 会报错（用了 luatexja）
2. 封面由学院统一发放，本模板提供的是内封面信息模板
3. 版权使用授权书需从研究生院网站下载替换
4. 目录由 LaTeX 自动生成，无需手动编写
5. 如使用 `references.bib`，取消 main.tex 中 `\bibliography{references}` 的注释
6. 首次编译可能较慢（需加载字体），后续编译会快很多
