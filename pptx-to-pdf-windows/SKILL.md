---
name: pptx-to-pdf-windows
description: Windows 上用本机 PowerPoint COM 把 PPTX 转为 PDF 并压缩体积。当用户要求 PPT 转 PDF、转格式、压缩 PPT 导出时使用。包含 PrintRange=None、InsertFromFile 重建等关键坑位。
agent_created: true
---

# PPTX → PDF 转换与压缩（Windows + 本机 PowerPoint）

## 环境准备
- 检测工具：优先本机 PowerPoint COM（`POWERPNT.EXE`，Office16）；LibreOffice/Ghostscript 通常不存在。
- Python：用受管 venv（路径形如 `C:/Users/<user>/.workbuddy/binaries/python/envs/default`，按你的实际环境调整），依赖：`pywin32 pikepdf pillow pymupdf`。

## 关键坑位（必读）
1. **不要用 PowerShell 调 PowerPoint 的 `ExportAsFixedFormat`**——PS 的 COM 绑定必然失败（DISP_E_TYPEMISMATCH / "无法将类型 X 的值转换为类型 Object"）。改用 Python pywin32 早绑定（`gencache.EnsureDispatch("PowerPoint.Application")`）。
2. **`ExportAsFixedFormat` 的第 8 个参数 `PrintRange` 是 VT_DISPATCH（COM 对象）类型，必须显式传 `None`**，否则报 "The Python instance can not be converted to a COM object"。
3. pywin32 动态绑定（DispatchEx）会因 makepy 默认值踩坑 2，务必用 `gencache.EnsureDispatch`（早绑定）并按关键字传参。
4. 导出/SaveAs 报 "PowerPoint 储存此文件时发生错误"（且 SaveCopyAs 也失败）时，该文件无法被 PowerPoint 保存。两条绕过路径，优先试 A：
   - **A（更可靠，首选）剥离内嵌字体**：见坑位 7。绝大多数"储存错误"的真正元凶是超大内嵌字体（几十 MB 的 `ppt/fonts/*.fntdata`）。
   - **B `InsertFromFile` 重建**：新建空演示文稿，`newPres.Slides.InsertFromFile(SRC, idx, start, end)` 分批（每批 5 页）插入全部页面，再导出。**注意：本方法在部分环境/文件上会直接失效**——报 `Integer out of range. 1 is not in FileName's valid range of 0 to 0`（连新存的 1 页小文件都报同样错，说明是环境性问题而非文件问题）。此时必须改用 A。
5. PowerShell 工具经常不回传 stdout：脚本里把结果写日志文件，再用 Read 读日志。**ps1 脚本中不能含 `Remove-Item` 等删除命令**（会被安全钩子整体拦截、exit 1）。
6. 文件名含中文/「」时，不要把路径硬编码进 ps1（编码会被破坏导致"找不到路径"）；用 Python（UTF-8 源文件没问题）或先复制到 ASCII 路径。
6b. **PowerPoint COM 的 `Presentations.Open` 对含空格的长路径会直接报 `com_error 0x80070003`（"发生意外"/路径找不到）**——即使文件确实存在。工作目录必须是无空格的短 ASCII 路径（如 `C:\pptwork\s.pptx`），并一律用反斜杠传给 COM；源目录形如 `E:\2. My Documents...\` 带空格时先复制出来再打开。
7. **剥离内嵌字体（解决"储存错误"的核心手段）**：用 Python `zipfile` 重打包，删 4 处引用，缺一不可，否则 PowerPoint 打不开：
   - 删除 `ppt/fonts/*.fntdata`（及 `.odttf/.ttf/.otf`）实体文件；
   - `ppt/_rels/presentation.xml.rels`：删 `Type=".../relationships/font"`（**注意是 `/font` 不是 `/fontData`**）的 `<Relationship>`；
   - `ppt/presentation.xml`：删 `<p:embeddedFontLst>…</p:embeddedFontLst>` 整块（以及残留的 `<p:embeddedFont .../>`）——**漏掉这步会因 r:id 悬空而打不开**；
   - `[Content_Types].xml`：删 `PartName="…/fonts/…"` 的 `<Override>`。
   剥离后 PowerPoint 用系统字体替代，导出即成功。体积也顺带暴降（本例 100MB→19.8MB）。
8. **不要试图用 pikepdf/PyMuPDF 二次压缩 PowerPoint 导出的 PDF**：降幅 <2%（甚至变大）；也不要"渲染成图片重建 PDF"——反而变大（PowerPoint Screen intent 的原生降采样+JPEG 已经接近最优）。真正有效的压缩手段只有 `Intent=1(Screen)`。PyMuPDF 存图参数名是 `pix.tobytes("jpeg", jpg_quality=...)`。

## 标准流程
1. 复制源文件到 ASCII 工作目录（如 `<workspace>/work/src.pptx`）。
2. `gencache.EnsureDispatch` 打开（`app.Visible=-1`，`WithWindow=True`）。
3. 直接尝试 `pres.ExportAsFixedFormat(dst, 2, 1, 0, 1, 1, 0, None, 1)`：
   - 参数：`FixedFormatType=2`(PDF)，`Intent`：1=Screen(小体积/降采样~46%)，2=Print(保原分辨率)。建议**两个 Intent 各导一份**，最终交付压缩版(Screen)+ 高清版(Print)。
   - 报"储存错误"→ 先做**坑位 7 的剥离内嵌字体**（首选），再重试导出；仍失败才试 InsertFromFile（坑位 4 B，可能失效）。
4. 二次压缩基本无效（坑位 8）：体积大头已在 Intent=1 时解决，不要再跑 pikepdf/重渲染。
5. 质检：`pymupdf` 渲染 2-4 页为 PNG 后用 Read 看图；pikepdf 检查字体是否子集嵌入（含 /DescendantFonts 递归检查）；核对页数与隐藏页（`<p:sld show="0"`）。
6. 把最终 PDF 复制回源目录（与源文件同名、.pdf 后缀），`present_files` 交付。

## 参考：典型体积
- 96MB 课件的 34 页：Print intent ≈ 4.2MB（原分辨率），Screen intent ≈ 1.5MB（约 46% 降采样）。
- 100MB（含 79.8MB 内嵌字体）、34 页、16:9(1920×1080pt) 课件，剥离字体后：Screen ≈ 4.7MB，Print ≈ 10.6MB（页面为图片型，压缩空间有限）。
- 交付建议：默认给 Screen 压缩版 + Print 高清版两份；命名区分（如 `-压缩版.pdf` / `-高清版.pdf`），避免覆盖源目录里已存在的同名 PDF。
