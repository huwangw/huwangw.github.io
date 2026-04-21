const cheerio = require('cheerio');
const fs = require('fs-extra');
const path = require('path');

// -------------------------- 配置参数 --------------------------
const INPUT_DIR = './old_html_files';   // 待处理文件目录
const OUTPUT_DIR = './new_html_files';  // 输出目录
const FILE_EXTENSIONS = ['.html', '.aspx'];  // 处理的文件后缀
// 过期属性映射规则：{原属性: [CSS属性, 默认单位]}
const ATTR_MAPPING = {
  align: ['text-align', ''],         // text-align无需单位
  bgcolor: ['background-color', ''], // 颜色值无需单位
  border: ['border', 'px solid'],    // border需要单位+样式
  cellspacing: ['border-spacing', 'px'],
  cellpadding: ['padding', 'px'],
  width: ['width', 'px'],
  height: ['height', 'px']
};
// --------------------------------------------------------------

/**
 * 为纯数字值添加单位（如100 → 100px，已带单位则不变）
 * @param {string} value - 属性值
 * @param {string} unit - 默认单位（如px）
 * @returns {string} 带单位的属性值
 */
function addUnit(value, unit) {
  if (!unit) return value; // 无需单位的属性直接返回
  // 匹配纯数字（整数/小数）
  const numRegex = /^\d+(\.\d+)?$/;
  if (numRegex.test(value.trim())) {
    return `${value}${unit}`;
  }
  return value; // 已带单位或非数字，直接返回
}

/**
 * 解析原有style属性为对象（如"color: red; width: 100" → {color: 'red', width: '100'}）
 * @param {string} styleStr - 原有style字符串
 * @returns {object} 解析后的样式对象
 */
function parseStyle(styleStr) {
  const styleObj = {};
  if (!styleStr) return styleObj;
  // 按分号分割样式，过滤空值
  const styleParts = styleStr.split(';').map(part => part.trim()).filter(Boolean);
  styleParts.forEach(part => {
    const [key, value] = part.split(':').map(item => item.trim());
    if (key && value) {
      styleObj[key] = value;
    }
  });
  return styleObj;
}

/**
 * 将样式对象转换为style字符串（如{color: 'red'} → "color: red;"）
 * @param {object} styleObj - 样式对象
 * @returns {string} style字符串
 */
function stringifyStyle(styleObj) {
  return Object.entries(styleObj)
    .map(([key, value]) => `${key}: ${value}`)
    .join('; ') + ';';
}

/**
 * 处理单个标签，替换过期属性为style
 * @param {CheerioElement} tag - 标签对象
 * @param {Cheerio} $ - cheerio实例
 */
function processTag(tag, $) {
  const $tag = $(tag);
  const styleObj = {};
  const attrsToRemove = [];

  // 处理过期属性，转换为style
  Object.keys(ATTR_MAPPING).forEach(attr => {
    if ($tag.attr(attr) !== undefined) {
      const [cssProp, unit] = ATTR_MAPPING[attr];
      const originalValue = $tag.attr(attr);
      const styledValue = addUnit(originalValue, unit);
      styleObj[cssProp] = styledValue; // 新样式覆盖旧样式
      attrsToRemove.push(attr); // 标记需要删除的原属性
    }
  });

  // 合并原有style属性
  const existingStyle = $tag.attr('style') || '';
  const existingStyleObj = parseStyle(existingStyle);
  // 原有样式中未被新样式覆盖的部分保留
  Object.keys(existingStyleObj).forEach(key => {
    if (!styleObj[key]) {
      styleObj[key] = existingStyleObj[key];
    }
  });

  // 更新标签的style属性
  if (Object.keys(styleObj).length > 0) {
    $tag.attr('style', stringifyStyle(styleObj));
  }

  // 删除原有的过期属性
  attrsToRemove.forEach(attr => {
    $tag.removeAttr(attr);
  });
}

/**
 * 处理单个文件
 * @param {string} inputPath - 输入文件路径
 * @param {string} outputPath - 输出文件路径
 */
async function processFile(inputPath, outputPath) {
  try {
    // 读取文件内容（UTF-8编码）
    const content = await fs.readFile(inputPath, 'utf8');
    // 解析HTML
    const $ = cheerio.load(content, {
      decodeEntities: false, // 保留原始字符（如&nbps;）
      xmlMode: false // 按HTML模式解析
    });

    // 遍历所有标签并处理
    $('*').each((i, tag) => {
      processTag(tag, $);
    });

    // 确保输出目录存在
    await fs.ensureDir(path.dirname(outputPath));
    // 保存处理后的内容
    await fs.writeFile(outputPath, $.html(), 'utf8');

    console.log(`处理完成：${inputPath} → ${outputPath}`);
  } catch (err) {
    console.error(`处理失败 ${inputPath}：${err.message}`);
  }
}

/**
 * 批量处理目录下的所有文件
 */
async function batchProcess() {
  try {
    // 遍历输入目录
    const walk = async (dir) => {
      const entries = await fs.readdir(dir, { withFileTypes: true });
      for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          await walk(fullPath); // 递归处理子目录
        } else if (entry.isFile()) {
          // 只处理指定后缀的文件
          const ext = path.extname(entry.name).toLowerCase();
          if (FILE_EXTENSIONS.includes(ext)) {
            // 计算输出路径（保持目录结构）
            const relativePath = path.relative(INPUT_DIR, fullPath);
            const outputPath = path.join(OUTPUT_DIR, relativePath);
            await processFile(fullPath, outputPath);
          }
        }
      }
    };

    await walk(INPUT_DIR);
    console.log('所有文件处理完毕！');
  } catch (err) {
    console.error(`批量处理失败：${err.message}`);
  }
}

// 启动处理
batchProcess();












{
    // 开启保存自动修复（核心）
    "editor.codeActionsOnSave": {
        "source.fixAll.htmlCodeCleaner": true
    },

    // ========== HTML Code Cleaner 专属配置 ==========
    // 开启：转换废弃属性 → 内联style（align/width/height等）
    "htmlCodeCleaner.convertDeprecatedAttributesToStyle": true,
    // 转换后删除原来的旧属性（必须开）
    "htmlCodeCleaner.removeOriginalDeprecatedAttrs": true,
    // 自动补px单位（width/height数值自动加px）
    "htmlCodeCleaner.autoAddPixelUnit": true,

    // 关闭所有多余清理，**只改废弃属性，不动class/id/其他代码**
    "htmlCodeCleaner.removeEmptyTags": false,
    "htmlCodeCleaner.removeHtmlComments": false,
    "htmlCodeCleaner.minifyCode": false,
    "htmlCodeCleaner.formatHtml": false,
    "htmlCodeCleaner.removeUnnecessarySpaces": false
}










{
  "extends": ["html-validate:recommended"],
  "rules": {
    // 开启：检测所有废弃HTML属性（align/width/height/bgcolor等）
    "no-deprecated-attr": "error",
    // 开启自动修复能力
    "no-deprecated-attr/fixable": true,
    // 关闭其他你不需要的冗余规则，避免干扰
    "doctype-html": "off",
    "no-unused-id": "off",
    "attribute-allowed-values": "warn"
  }
}










<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>原生HTML 转 @material/web 转换器【无依赖·全兼容换行Input】</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: system-ui, -apple-system, sans-serif;
        }
        body {
            padding: 24px;
            max-width: 1400px;
            margin: 0 auto;
            background: #f5f5f5;
        }
        .title {
            margin-bottom: 12px;
            font-size: 18px;
            font-weight: 600;
            color: #333;
        }
        .container {
            display: flex;
            gap: 20px;
            margin-bottom: 16px;
        }
        .box {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        textarea {
            width: 100%;
            height: 420px;
            padding: 14px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 14px;
            line-height: 1.7;
            background: #fff;
            resize: vertical;
            white-space: pre;
        }
        .btn-group {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
        }
        button {
            padding: 9px 22px;
            border: none;
            border-radius: 6px;
            background: #333;
            color: #fff;
            font-size: 14px;
            cursor: pointer;
        }
        button:hover {
            background: #555;
        }
        .tip {
            font-size: 12px;
            color: #666;
            margin-top: 6px;
            line-height: 1.5;
        }
    </style>
</head>
<body>
    <div class="title">原始普通 HTML（粘贴原生代码）</div>
    <div class="container">
        <div class="box">
            <textarea id="source" placeholder="粘贴任意原生HTML，input无论多少属性、怎么换行、跨行都可识别转换"></textarea>
        </div>
        <div class="box">
            <textarea id="result" placeholder="转换完成的 @material/web 代码"></textarea>
        </div>
    </div>

    <div class="btn-group">
        <button onclick="convert()">一键转换为 material/web</button>
        <button onclick="clearAll()">清空全部</button>
    </div>

    <div class="tip">
        ✅ 纯原生JS，无任何外部依赖、无CDN、离线可用<br>
        ✅ 完美兼容：input属性过多、标签换行、跨行、多属性乱序全部可识别<br>
        ✅ 新增：input type="button" → md-outlined-button<br>
        ✅ 完整input类型映射，所有属性(id/class/name/value等)全部原样保留<br>
        ✅ 基于原生DOMParser解析，不再使用不稳定正则匹配
    </div>

<script>
function convert() {
    const sourceHtml = document.getElementById('source').value.trim();
    if (!sourceHtml) {
        alert('请粘贴原始HTML代码');
        return;
    }

    // 原生DOM解析，无视所有换行、空格、格式
    const parser = new DOMParser();
    const doc = parser.parseFromString(sourceHtml, 'text/html');
    const container = doc.body;

    // ========== 完整映射表 ==========
    const inputMap = {
        text: 'md-outlined-text-field',
        password: 'md-outlined-text-field',
        number: 'md-outlined-text-field',
        checkbox: 'md-checkbox',
        radio: 'md-radio',
        range: 'md-slider',
        // 你新增的：input type="button" 转为轮廓按钮
        button: 'md-outlined-button'
    };

    // 遍历所有input标签，全部转换
    const allInputs = container.querySelectorAll('input');
    allInputs.forEach(oldInput => {
        const type = oldInput.type || 'text';
        const newTag = inputMap[type];
        if (!newTag) return;

        const newEl = document.createElement(newTag);
        // 复制全部所有属性，一个不漏
        copyAllAttr(oldInput, newEl);
        oldInput.parentNode.replaceChild(newEl, oldInput);
    });

    // 原生 <button> 标签 → md-button
    container.querySelectorAll('button').forEach(old => {
        const el = document.createElement('md-button');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // <a> 链接 → md-link
    container.querySelectorAll('a').forEach(old => {
        const el = document.createElement('md-link');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // <textarea> → md-textarea
    container.querySelectorAll('textarea').forEach(old => {
        const el = document.createElement('md-textarea');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // <select> → md-select
    container.querySelectorAll('select').forEach(old => {
        const el = document.createElement('md-select');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // <option> → md-option
    container.querySelectorAll('option').forEach(old => {
        const el = document.createElement('md-option');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // div class="card" → md-card
    container.querySelectorAll('div.card').forEach(old => {
        const el = document.createElement('md-card');
        el.innerHTML = old.innerHTML;
        copyAllAttr(old, el);
        old.replaceWith(el);
    });

    // 导出最终HTML代码
    const resultHtml = container.innerHTML;
    document.getElementById('result').value = resultHtml;
}

// 通用工具：复制元素全部属性（包含自定义data-*属性）
function copyAllAttr(source, target) {
    Array.from(source.attributes).forEach(attr => {
        target.setAttribute(attr.name, attr.value);
    });
}

// 清空全部
function clearAll() {
    document.getElementById('source').value = '';
    document.getElementById('result').value = '';
}
</script>
</body>
</html>

















<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>原生HTML 转 @material/web 转换器【纯本地无依赖】</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: system-ui, -apple-system, sans-serif;
        }
        body {
            padding: 24px;
            max-width: 1400px;
            margin: 0 auto;
            background: #f5f5f5;
        }
        .title {
            margin-bottom: 12px;
            font-size: 18px;
            font-weight: 600;
            color: #333;
        }
        .container {
            display: flex;
            gap: 20px;
            margin-bottom: 16px;
        }
        .box {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        textarea {
            width: 100%;
            height: 420px;
            padding: 14px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 14px;
            line-height: 1.7;
            background: #fff;
            resize: vertical;
            white-space: pre;
        }
        .btn-group {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
        }
        button {
            padding: 9px 22px;
            border: none;
            border-radius: 6px;
            background: #333;
            color: #fff;
            font-size: 14px;
            cursor: pointer;
        }
        button:hover {
            background: #555;
        }
        .tip {
            font-size: 12px;
            color: #666;
            margin-top: 6px;
        }
    </style>
</head>
<body>
    <div class="title">原始普通 HTML（粘贴原生代码）</div>
    <div class="container">
        <div class="box">
            <textarea id="source" placeholder="在此粘贴原生HTML：button、input、select、a、div、textarea、checkbox、radio 等全部原生标签"></textarea>
        </div>
        <div class="box">
            <textarea id="result" placeholder="转换完成的 @material/web 代码会在此显示，直接复制使用"></textarea>
        </div>
    </div>

    <div class="btn-group">
        <button onclick="convert()">一键转换为 material/web 标签</button>
        <button onclick="clearAll()">清空全部内容</button>
    </div>

    <div class="tip">
        ✅ 本工具纯本地原生JS编写，无任何外部CDN、无第三方包、无外部依赖、无网络请求<br>
        ✅ 仅做标签替换转换，保留所有原生属性（id/class/name/value/placeholder/disabled/required等）<br>
        ✅ 转换后代码直接复制到你的项目，使用你项目自身已引入的 @material/web 即可运行
    </div>

<script>
/**
 * 纯原生JS转换核心
 * 无任何外部依赖、无外部脚本、无网络请求
 * 原生HTML 自动转为 @material/web M3 自定义标签
 */
function convert() {
    let html = document.getElementById('source').value;
    if (!html.trim()) {
        alert('请先粘贴原始HTML代码');
        return;
    }

    // 全部转换规则：正则 + 替换
    // 统一非贪婪匹配，兼容：多空格、单/双引号、自闭合标签、无闭合标签、属性乱序
    const rules = [
        // 1. 原生 button -> md-button
        [/<button([^>]*?)>([\s\S]*?)<\/button>/gi, '<md-button$1>$2</md-button>'],

        // 2. 文本输入框 input type="text"
        [/<input([^>]*?)type\s*=\s*(["'])text\1([^>]*?)\/?>/gi, '<md-outlined-text-field$1$3></md-outlined-text-field>'],
        [/<input([^>]*?)type\s*=\s*(["'])text\1([^>]*?)>/gi, '<md-outlined-text-field$1$3></md-outlined-text-field>'],

        // 3. 密码框 password
        [/<input([^>]*?)type\s*=\s*(["'])password\1([^>]*?)\/?>/gi, '<md-outlined-text-field type="password"$1$3></md-outlined-text-field>'],
        [/<input([^>]*?)type\s*=\s*(["'])password\1([^>]*?)>/gi, '<md-outlined-text-field type="password"$1$3></md-outlined-text-field>'],

        // 4. 数字输入 number
        [/<input([^>]*?)type\s*=\s*(["'])number\1([^>]*?)\/?>/gi, '<md-outlined-text-field type="number"$1$3></md-outlined-text-field>'],
        [/<input([^>]*?)type\s*=\s*(["'])number\1([^>]*?)>/gi, '<md-outlined-text-field type="number"$1$3></md-outlined-text-field>'],

        // 5. 复选框 checkbox
        [/<input([^>]*?)type\s*=\s*(["'])checkbox\1([^>]*?)\/?>/gi, '<md-checkbox$1$3></md-checkbox>'],
        [/<input([^>]*?)type\s*=\s*(["'])checkbox\1([^>]*?)>/gi, '<md-checkbox$1$3></md-checkbox>'],

        // 6. 单选框 radio
        [/<input([^>]*?)type\s*=\s*(["'])radio\1([^>]*?)\/?>/gi, '<md-radio$1$3></md-radio>'],
        [/<input([^>]*?)type\s*=\s*(["'])radio\1([^>]*?)>/gi, '<md-radio$1$3></md-radio>'],

        // 7. 开关 switch
        [/<input([^>]*?)type\s*=\s*(["'])checkbox\1([^>]*?)class\s*=\s*(["'])switch\4([^>]*?)\/?>/gi, '<md-switch$1$5></md-switch>'],
        [/<input([^>]*?)type\s*=\s*(["'])checkbox\1([^>]*?)class\s*=\s*(["'])switch\4([^>]*?)>/gi, '<md-switch$1$5></md-switch>'],

        // 8. 超链接 a -> md-link
        [/<a([^>]*?)>([\s\S]*?)<\/a>/gi, '<md-link$1>$2</md-link>'],

        // 9. 文本域 textarea
        [/<textarea([^>]*?)>([\s\S]*?)<\/textarea>/gi, '<md-textarea$1>$2</md-textarea>'],

        // 10. 下拉框 select + option
        [/<select([^>]*?)>([\s\S]*?)<\/select>/gi, '<md-select$1>$2</md-select>'],
        [/<option([^>]*?)>([\s\S]*?)<\/option>/gi, '<md-option$1>$2</md-option>'],

        // 11. 卡片 div class="card"
        [/<div([^>]*?)class\s*=\s*(["'])card\2([^>]*?)>([\s\S]*?)<\/div>/gi, '<md-card$1$3>$4</md-card>'],

        // 12. 滑块 range
        [/<input([^>]*?)type\s*=\s*(["'])range\1([^>]*?)\/?>/gi, '<md-slider$1$3></md-slider>'],
        [/<input([^>]*?)type\s*=\s*(["'])range\1([^>]*?)>/gi, '<md-slider$1$3></md-slider>'],
    ];

    // 依次执行所有转换规则
    rules.forEach(([reg, replace]) => {
        html = html.replace(reg, replace);
    });

    document.getElementById('result').value = html;
}

// 清空所有输入输出
function clearAll() {
    document.getElementById('source').value = '';
    document.getElementById('result').value = '';
}
</script>
</body>
</html>














{
  // 关闭vscode原生自带弱CSS校验，避免冲突
  "css.validate": false,
  "less.validate": false,
  "scss.validate": false,

  // Stylelint 核心配置：开启、实时校验、保存自动修复
  "stylelint.validate": ["css", "less", "scss", "html"],
  "editor.codeActionsOnSave": {
    // 保存自动修复所有Stylelint错误（包含缺px自动补全）
    "source.fixAll.stylelint": true
  },

  // 强制所有宽高尺寸必须带px单位，无单位直接报错+自动补px
  "stylelint.config": {
    "rules": {
      // 核心规则：禁止尺寸值无单位（width/height/top/left等）
      "length-no-zero-units": null,
      "unit-no-unknown": true,
      "value-no-vendor-prefix": null,
      // 自定义：数字必须带px，0可以无单位（0 不用px）
      "declaration-property-value-allowed-list": {
        "/^(width|height|top|left|right|bottom|margin|padding)/": ["/^\\d+px$/", "0"]
      }
    }
  }
}













83import com.aspose.cells.*;

public void setupColumnDataValidations(Worksheet worksheet, int strtRow, it endRow) {
    DataValidationCollection validations = worksheet.getValidations();
    
    // 建议先清除目标区域旧的验证，防止冲突（这是解决你报错的关键）
    // 清除 A 列和 C 列在目标数据区域的旧验证
    validations.clear(startRow, 0, endRow - startRow + 1, 1); // A列
    validations.clear(startRow, 2, endRow - startRow + 1, 1); // C列

    // 1. 设置 A 列的验证 (例如：引用另一个Sheet的命名区域 "StatusList")
    addListValidation(validations, 0, startRow, endRow, "=StatusList");

    // 2. 设置 C 列的验证 (例如：引用另一个Sheet的具体范围)
    // 注意：如果是固定范围，可以直接写；如果是动态命名区域，建议也定义成名字
    addListValidation(validations, 2, startRow, endRow, "=DataSource!$A$1:$A$50");
    
    // 可选：关闭自动更新引用以提高性能
    // validations.setUpdateReference(false);
}

/**
 * 辅助方法：为指定列添加下拉列表验证
 */
private void addListValidation(DataValidationCollection validations, 
                               int column, int startRow, int endRow, String formula) {
    DataValidation validation = validations.add();
    
    // 设置区域：整列范围
    CellArea area = new CellArea();
    area.StartRow = startRow;
    area.EndRow = endRow;
    area.StartColumn = column;
    area.EndColumn = column;
    validation.setArea(area);
    
    // 设置为列表类型
    validation.setType(ValidationType.LIST);
    validation.setFormula1(formula); // 这里填入命名区域名或跨Sheet引用
    validation.setShowError(true);
    validation.setInCellDropdown(true); // 显示下拉箭头
}





import com.aspose.cells.*;

import java.util.ArrayList;
import java.util.List;

public class DataValidationUtil {

    /**
     * 获取指定范围内的所有数据验证规则
     * @param worksheet 要搜索的工作表
     * @param targetRange 目标范围 (例如: A10:D1000)
     * @return 包含规则及其在范围内具体位置的列表
     */
    public static List<ValidationInfo> getValidationsInRange(Worksheet worksheet, Range targetRange) {
        List<ValidationInfo> result = new ArrayList<>();
        DataValidationCollection validations = worksheet.getValidations();

        // 获取目标范围的行列边界
        int tgtRowStart = targetRange.getFirstRow();
        int tgtRowEnd = targetRange.getFirstRow() + targetRange.getRowCount() - 1;
        int tgtColStart = targetRange.getFirstColumn();
        int tgtColEnd = targetRange.getFirstColumn() + targetRange.getColumnCount() - 1;

        // 遍历工作表上所有的数据验证规则
        for (int i = 0; i < validations.getCount(); i++) {
            DataValidation validation = validations.get(i);
            CellArea area = validation.getArea();

            // 检查该规则的区域是否与目标范围有重叠 (Intersect)
            if (isIntersect(area, tgtRowStart, tgtRowEnd, tgtColStart, tgtColEnd)) {
                // 创建一个包含规则和重叠区域信息的对象
                ValidationInfo info = new ValidationInfo();
                info.validation = validation;
                info.appliedArea = area; // 原始应用区域
                // 计算实际在目标范围内的区域 (可选，用于精确重建)
                info.effectiveArea = calculateIntersection(area, tgtRowStart, tgtRowEnd, tgtColStart, tgtColEnd);
                result.add(info);
            }
        }
        return result;
    }

    // 检查两个区域是否有交集
    private static boolean isIntersect(CellArea area, int r1, int r2, int c1, int c2) {
        return !(area.EndRow < r1 || area.StartRow > r2 || area.EndColumn < c1 || area.StartColumn > c2);
    }

    // 计算交集区域
    private static CellArea calculateIntersection(CellArea area, int r1, int r2, int c1, int c2) {
        CellArea intersection = new CellArea();
        intersection.StartRow = Math.max(area.StartRow, r1);
        intersection.EndRow = Math.min(area.EndRow, r2);
        intersection.StartColumn = Math.max(area.StartColumn, c1);
        intersection.EndColumn = Math.min(area.EndColumn, c2);
        return intersection;
    }

    // 包装类：用于存储验证规则及其在特定范围内的信息
    public static class ValidationInfo {
        public DataValidation validation;
        public CellArea appliedArea;   // 规则原本应用的区域
        public CellArea effectiveArea; // 规则在目标范围内的实际生效区域
    }
}





// 1. 定义你复制的目标范围
Range targetRange = worksheet.getCells().createRange("A10:C5000");

// 2. 获取该范围内的所有验证规则
List<DataValidationUectiveArea.StartColumn; // 获取到的列
    newArea.EndColumn = info.effectiveArea.EndColumn;
    newDV.setArea(newArea);
}







import com.aspose.cells.*;

import java.util.*;

public class TemplateValidationAnalyzer {

    // 用于存储分析结果的内部类
    public static class ColumnValidationRule {
        public int columnIndex;
        public String columnName;
        public String validationType;
        public String formula1;
        public String formula2; // 用于介于、大于等于等双值条件
        public String operator;

        @Override
        public String toString() {
            return String.format("列: %s (Index: %d), 类型: %s, 操作符: %s, 公式1: %s, 公式2: %s",
                    columnName, columnIndex, validationType, operator, formula1, formula2);
        }
    }

    /**
     * 分析指定工作表，获取所有使用了数据验证的列及其规则
     * @param worksheet 模板工作表
     * @return 包含列验证规则的列表
     */
    public static List<ColumnValidationRule> analyzeValidationsInSheet(Worksheet worksheet) {
        List<ColumnValidationRule> result = new ArrayList<>();
        ValidationCollection validations = worksheet.getValidations();

        // 遍历工作表上所有的验证规则
        for (int i = 0; i < validations.getCount(); i++) {
            DataValidation validation = validations.get(i);
            
            // *** 关键：获取该规则应用的所有区域 ***
            // 一个验证规则可能应用在不连续
            
            
            
            
            
            
            
            
            
            
            
            
            
            constst readline = require('readline');

/**
 * 核心转换函数：将 LINQ 字符串解析为 SQL
 * @param {string} linqQuery - 用户输入的 LINQ 语句
 * @returns {string} - 生成的 SQL 语句
 */
function parseLinqToSql(linqQuery) {
    // 1. 预处理：去除首尾空格，统一换行符
    let query = linqQuery.trim().replace(/\r\n/g, ' ').replace(/\n/g, ' ');

    // --- 步骤 A: 解析 FROM 子句 ---
    // 匹配模式: from [别名] in context.[表名]
    // 例如: from re in context.MT
    const fromRegex = /from\s+(\w+)\s+in\s+context\.(\w+)/i;
    const fromMatch = query.match(fromRegex);

    if (!fromMatch) {
        return "❌ 错误：无法识别 'from' 语句。请确保格式为 'from 别名 in context.表名'";
    }

    const alias = fromMatch[1];      // 获取别名，例如 "re"
    const tableName = fromMatch[2];  // 获取表名，例如 "MT"

    // --- 步骤 B: 解析 WHERE 子句 ---
    // 匹配模式: where [别名].[列名].equals([值])
    // 注意：这里使用了动态构建的正则，因为别名是上一步解析出来的
    // 例如: where re.CODE.equals(code)
    const whereRegex = new RegExp(`where\\s+${alias}\\.(\\w+)\\.equals\$([^)]+)\$`, 'i');
    const whereMatch = query.match(whereRegex);

    let whereClause = "";
    if (whereMatch) {
        const columnName = whereMatch[1]; // 例如 "CODE"
        const paramValue = whereMatch[2].trim(); // 例如 "code"
        // 生成参数化查询占位符
        whereClause = `WHERE [${columnName}] = @${paramValue}`;
    }

    // --- 步骤 C: 解析 SELECT 子句 ---
    // 匹配模式: select new [对象]{ [属性] = [别名].[列], ... }
    // 例如: select new Md(){ Code= re.code, Name = re.name }
    const selectRegex = new RegExp(`select\\s+new\\s+\\w+\\s*\$\\s*\$\\s*\\{([^}]+)\\}`, 'i');
    const selectMatch = query.match(selectRegex);

    let selectClause = "*"; // 默认查所有
    if (selectMatch) {
        const selectBody = selectMatch[1]; // 获取花括号内的内容
        
        // 使用正则查找所有 "别名.列名" 的模式
        // 匹配结果如: ["re.code", "re.name"]
        const columns = selectBody.match(new RegExp(`${alias}\\.(\\w+)`, 'g'));
        
        if (columns) {
            // 提取列名并加上中括号 [Code], [Name]
            const formattedCols = columns.map(col => {
                const colName = col.split('.')[1];
                return `[${colName}]`;
            });
            selectClause = formattedCols.join(', ');
        }
    }

    // --- 步骤 D: 组装最终 SQL ---
    let sql = `SELECT ${selectClause}\nFROM [${tableName}]`;
    if (whereClause) {
        sql += `\n${whereClause};`;
    } else {
        sql += `;`;
    }

    return sql;
}

// --- 命令行交互界面 ---
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

console.log('------------------------------------------------');
console.log('🚀  LINQ to SQL 转换工具 (Node.js 版)');
console.log('💡  支持语法: from x in context.T where x.C.equals(v) select new O(){...}');
console.log('输入 "exit" 或 "quit" 退出程序');
console.log('------------------------------------------------');

function startPrompt() {
    rl.question('\n👉 请输入 LINQ 语句:\n> ', (input) => {
        if (input.toLowerCase() === 'exit' || input.toLowerCase() === 'quit') {
            console.log('👋 再见！');
            rl.close();
            return;
        }

        try {
            const result = parseLinqToSql(input);
            console.log('\n✅ 生成的 SQL:');
            console.log('------------------------------------------------');
            console.log(result);
            console.log('------------------------------------------------');
        } catch (err) {
            console.log('❌ 发生未知错误:', err);
        }

        startPrompt(); // 循环继续询问
    });
}

startPrompt();









const readline = require('readline');

/**
 * 终极版 LINQ 转 SQL 解析器
 * 支持：From, Where, Join (Inner/Left), Select New, Null-Coalescing (??), Ternary (?:)
 */
function parseLinqToSql(linqQuery) {
    // 1. 预处理：清洗字符串
    let query = linqQuery.trim().replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\s+/g, ' ');

    // --- 步骤 A: 解析 FROM 子句 ---
    const fromRegex = /from\s+(\w+)\s+in\s+context\.(\w+)/i;
    const fromMatch = query.match(fromRegex);
    if (!fromMatch) return "❌ 错误：无法识别 'from' 语句。";
    const mainAlias = fromMatch[1];
    const mainTable = fromMatch[2];

    // --- 步骤 B: 解析 JOIN 子句 ---
    let joinClause = "";
    
    // Left Join (GroupJoin + DefaultIfEmpty)
    const leftJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)\s+into\s+(\w+).*?from\s+(\w+)\s+in\s+\7\.DefaultIfEmpty$$/gi;
    let match;
    while ((match = leftJoinRegex.exec(query)) !== null) {
        const jTable = match[2];
        const leftA = match[3]; const leftC = match[4];
        const rightA = match[5]; const rightC = match[6];
        const finalAlias = match[8];
        joinClause += `\nLEFT JOIN [${jTable}] AS [${finalAlias}] ON [${leftA}].[${leftC}] = [${rightA}].[${rightC}]`;
    }

    // Inner Join
    const innerJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)(?!\s+into)/gi;
    while ((match = innerJoinRegex.exec(query)) !== null) {
        const jTable = match[2];
        const leftA = match[3]; const leftC = match[4];
        const rightA = match[5]; const rightC = match[6];
        const jAlias = match[1];
        joinClause += `\nINNER JOIN [${jTable}] AS [${jAlias}] ON [${leftA}].[${leftC}] = [${rightA}].[${rightC}]`;
    }

    // --- 步骤 C: 解析 WHERE 子句 ---
    const whereStart = query.search(/where\s/i);
    const selectStart = query.search(/select\s/i);
    let whereClause = "";
    if (whereStart !== -1 && selectStart !== -1) {
        let whereBody = query.substring(whereStart + 6, selectStart).trim();
        whereBody = whereBody.replace(/(\w+)\.(\w+)\.equals$([^)]+)$/g, '[$1].[$2] = $3');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*(\w+)\.(\w+)/g, '[$1].[$2] = [$3].[$4]');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*([^=\s]+)/g, '[$1].[$2] = $3');
        if (whereBody) whereClause = `WHERE ${whereBody}`;
    }

    // --- 步骤 D: 解析 SELECT 子句 (核心升级部分) ---
    const selectRegex = /select\s+new\s+\w+\s*$\s*$\s*\{([^}]+)\}/i;
    const selectMatch = query.match(selectRegex);
    let selectClause = "*";

    if (selectMatch) {
        const selectBody = selectMatch[1];
        const columns = [];
        
        // 正则匹配每一个字段定义，例如：Code = re.Code, 或 Name = t.Name ?? "Unknown"
        // 这里我们按逗号分割，但要小心对象初始化中的逗号（这里简化处理，假设没有嵌套对象）
        const fields = selectBody.split(',');

        fields.forEach(field => {
            field = field.trim();
            if (!field) return;

            // 1. 处理 ?? (空合并运算符)
            // 匹配模式: Alias = Col ?? Value
            const nullCoalesceMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)\s*\?\?\s*(.+)/);
            if (nullCoalesceMatch) {
                const colAlias = nullCoalesceMatch[1];
                const tableAlias = nullCoalesceMatch[2];
                const colName = nullCoalesceMatch[3];
                let defaultVal = nullCoalesceMatch[4].trim();

                // 处理默认值格式
                if (defaultVal === 'string.Empty') defaultVal = "''";
                else if (!isNaN(defaultVal)) defaultVal = defaultVal; // 数字
                else if (defaultVal.startsWith('"') && defaultVal.endsWith('"')) defaultVal = `'${defaultVal.slice(1, -1)}'`; // 字符串字面量

                columns.push(`ISNULL([${tableAlias}].[${colName}], ${defaultVal}) AS [${colAlias}]`);
                return;
            }

            // 2. 处理三元运算符 ?:
            // 匹配模式: Alias = Cond ? Val1 : Val2
            const ternaryMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)\s*\?\s*(.+)\s*:\s*(.+)/);
            if (ternaryMatch) {
                const colAlias = ternaryMatch[1];
                const tableAlias = ternaryMatch[2];
                const colName = ternaryMatch[3];
                const trueVal = ternaryMatch[4].trim();
                const falseVal = ternaryMatch[5].trim();
                
                // 简单处理值
                const formatVal = (v) => {
                     if (v === 'true') return '1';
                     if (v === 'false') return '0';
                     if (v.startsWith('"')) return `'${v.slice(1, -1)}'`;
                     return v;
                };

                columns.push(`CASE WHEN [${tableAlias}].[${colName}] = 1 THEN ${formatVal(trueVal)} ELSE ${formatVal(falseVal)} END AS [${colAlias}]`);
                return;
            }

            // 3. 普通字段: Alias = Col
            const simpleMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)/);
            if (simpleMatch) {
                const colAlias = simpleMatch[1];
                const tableAlias = simpleMatch[2];
                const colName = simpleMatch[3];
                columns.push(`[${tableAlias}].[${colName}] AS [${colAlias}]`);
            }
        });

        if (columns.length > 0) {
            selectClause = columns.join(', ');
        }
    }

    // --- 步骤 E: 组装 SQL ---
    let sql = `SELECT ${selectClause}\nFROM [${mainTable}] AS [${mainAlias}]`;
    if (joinClause) sql += joinClause;
    if (whereClause) sql += `\n${whereClause};`;
    else sql += `;`;

    return sql;
}

// --- 命令行交互 ---
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
console.log('------------------------------------------------');
console.log('🚀  LINQ to SQL 终极版 (支持 ?? 和 ?:)');
console.log('------------------------------------------------');

function startPrompt() {
    rl.question('\n👉 请输入 LINQ 语句:\n> ', (input) => {
        if (input.toLowerCase() === 'exit') { rl.close(); return; }
        try {
            const result = parseLinqToSql(input);
            console.log('\n✅ 生成的 SQL:');
            console.log('------------------------------------------------');
            console.log(result);
            console.log('------------------------------------------------');
        } catch (err) { console.log('❌ 错误:', err); }
        startPrompt();
    });
}
startPrompt();






const readline = require('readline');

/**
 * 最终版 LINQ 转 SQL 解析器
 * 重点增强了 WHERE 子句的转换逻辑
 */
function parseLinqToSql(linqQuery) {
    // 1. 预处理
    let query = linqQuery.trim().replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\s+/g, ' ');

    // --- 步骤 A: 解析 FROM ---
    const fromRegex = /from\s+(\w+)\s+in\s+context\.(\w+)/i;
    const fromMatch = query.match(fromRegex);
    if (!fromMatch) return "❌ 错误：无法识别 'from' 语句。";
    const mainAlias = fromMatch[1];
    const mainTable = fromMatch[2];

    // --- 步骤 B: 解析 JOIN ---
    let joinClause = "";
    // Left Join
    const leftJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)\s+into\s+(\w+).*?from\s+(\w+)\s+in\s+\7\.DefaultIfEmpty$$/gi;
    let match;
    while ((match = leftJoinRegex.exec(query)) !== null) {
        const jTable = match[2];
        const finalAlias = match[8];
        joinClause += `\nLEFT JOIN [${jTable}] AS [${finalAlias}] ON [${match[3]}].[${match[4]}] = [${match[5]}].[${match[6]}]`;
    }
    // Inner Join
    const innerJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)(?!\s+into)/gi;
    while ((match = innerJoinRegex.exec(query)) !== null) {
        joinClause += `\nINNER JOIN [${match[2]}] AS [${match[1]}] ON [${match[3]}].[${match[4]}] = [${match[5]}].[${match[6]}]`;
    }

    // --- 步骤 C: 解析 WHERE (核心升级部分) ---
    const whereStart = query.search(/where\s/i);
    const selectStart = query.search(/select\s/i);
    let whereClause = "";

    if (whereStart !== -1 && selectStart !== -1) {
        let whereBody = query.substring(whereStart + 6, selectStart).trim();

        // 1. 处理 .Equals() 语法
        // 匹配: x.Prop.Equals(val)
        // 如果 val 是 null，转换为 IS NULL，否则转换为 =
        whereBody = whereBody.replace(/(\w+)\.(\w+)\.Equals$([^)]+)$/g, (fullMatch, alias, col, val) => {
            val = val.trim();
            if (val === 'null') {
                return `[${alias}].[${col}] IS NULL`;
            }
            return `[${alias}].[${col}] = ${val}`;
        });

        // 2. 处理 == 和 != 语法
        // 处理 == null
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*null/g, '[$1].[$2] IS NULL');
        // 处理 != null 或 <> null
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*(!=|<>)\s*null/g, '[$1].[$2] IS NOT NULL');
        // 处理普通 !=
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*!=\s*(\w+)\.(\w+)/g, '[$1].[$2] <> [$3].[$4]');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*!=\s*(.+)/g, '[$1].[$2] <> $3');
        
        // 处理普通 == (非 null, 非对象比较)
        // 这里做一个简单的假设：如果是 对象.属性 == 对象.属性，我们已经处理了。
        // 如果是 对象.属性 == 变量，这里简单转换
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*(\w+)\.(\w+)/g, '[$1].[$2] = [$3].[$4]');
        
        // 3. 处理 .Contains() -> LIKE
        whereBody = whereBody.replace(/(\w+)\.(\w+)\.Contains$([^)]+)$/g, '[$1].[$2] LIKE '%' + $3 + '%'');

        if (whereBody) whereClause = `WHERE ${whereBody}`;
    }

    // --- 步骤 D: 解析 SELECT ---
    const selectRegex = /select\s+new\s+\w+\s*$\s*$\s*\{([^}]+)\}/i;
    const selectMatch = query.match(selectRegex);
    let selectClause = "*";

    if (selectMatch) {
        const selectBody = selectMatch[1];
        const columns = [];
        const fields = selectBody.split(',');

        fields.forEach(field => {
            field = field.trim();
            if (!field) return;

            // 处理 ??
            const nullCoalesceMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)\s*\?\?\s*(.+)/);
            if (nullCoalesceMatch) {
                const colAlias = nullCoalesceMatch[1];
                const tableAlias = nullCoalesceMatch[2];
                const colName = nullCoalesceMatch[3];
                let defaultVal = nullCoalesceMatch[4].trim();
                if (defaultVal === 'string.Empty') defaultVal = "''";
                else if (defaultVal.startsWith('"') && defaultVal.endsWith('"')) defaultVal = `'${defaultVal.slice(1, -1)}'`;
                columns.push(`ISNULL([${tableAlias}].[${colName}], ${defaultVal}) AS [${colAlias}]`);
                return;
            }

            // 普通字段
            const simpleMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)/);
            if (simpleMatch) {
                columns.push(`[${simpleMatch[2]}].[${simpleMatch[3]}] AS [${simpleMatch[1]}]`);
            }
        });

        if (columns.length > 0) selectClause = columns.join(', ');
    }

    // --- 步骤 E: 组装 ---
    let sql = `SELECT ${selectClause}\nFROM [${mainTable}] AS [${mainAlias}]`;
    if (joinClause) sql += joinClause;
    if (whereClause) sql += `\n${whereClause};`;
    else sql += `;`;

    return sql;
}

// --- 命令行交互 ---
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
console.log('------------------------------------------------');
console.log('🚀  LINQ to SQL 最终版 (完美支持 WHERE 语法)');
console.log('------------------------------------------------');

function startPrompt() {
    rl.question('\n👉 请输入 LINQ 语句:\n> ', (input) => {
        if (input.toLowerCase() === 'exit') { rl.close(); return; }
        try {
            const result = parseLinqToSql(input);
            console.log('\n✅ 生成的 SQL:');
            console.log('------------------------------------------------');
            console.log(result);
            console.log('------------------------------------------------');
        } catch (err) { console.log('❌ 错误:', err); }
        startPrompt();
    });
}
startPrompt();


const readline = require('readline');

/**
 * 混合连接修复版 LINQ 转 SQL 解析器
 * 重点修复：同时存在 Inner Join 和 Left Join 时的解析错误
 */

// --- 辅助函数：智能转换 Equals 表达式 ---
function convertEqualsExpression(match, alias, col, valueContent) {
    let val = valueContent.trim();
    if (val === 'null') return `[${alias}].[${col}] IS NULL`;
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
        return `[${alias}].[${col}] = '${val.slice(1, -1)}'`;
    }
    if (!isNaN(val)) return `[${alias}].[${col}] = ${val}`;
    return `[${alias}].[${col}] = @${val}`;
}

function parseLinqToSql(linqQuery) {
    // 1. 预处理
    let query = linqQuery.trim().replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\s+/g, ' ');

    // --- 步骤 A: 解析 FROM ---
    const fromRegex = /from\s+(\w+)\s+in\s+context\.(\w+)/i;
    const fromMatch = query.match(fromRegex);
    if (!fromMatch) return "❌ 错误：无法识别 'from' 语句。";
    const mainAlias = fromMatch[1];
    const mainTable = fromMatch[2];

    // --- 步骤 B: 解析 JOIN (核心修复部分) ---
    let joinClause = "";

    // 1. 优先处理 LEFT JOIN (GroupJoin + DefaultIfEmpty)
    // 逻辑：join ... into [组别名] ... from [新别名] in [组别名].DefaultIfEmpty()
    const leftJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)\s+into\s+(\w+)\s+from\s+(\w+)\s+in\s+\7\.DefaultIfEmpty$$/gi;
    
    let match;
    while ((match = leftJoinRegex.exec(query)) !== null) {
        const jTable = match[2];      // 关联表名 (例如 M_F)
        const leftA = match[3];       // 左表别名
        const leftC = match[4];       // 左表列
        const rightA = match[5];      // 右表别名 (原始)
        const rightC = match[6];      // 右表列
        // const groupAlias = match[7]; // 分组别名 (中间变量，SQL不需要)
        const finalAlias = match[8];  // 最终展开的别名 (例如 fos)

        joinClause += `\nLEFT JOIN [${jTable}] AS [${finalAlias}] ON [${leftA}].[${leftC}] = [${rightA}].[${rightC}]`;
    }

    // 2. 处理 INNER JOIN (标准 Join)
    // 逻辑：join ... on ... equals ... (且后面不能紧跟 into)
    // 注意：这里使用了负向预查 (?!\s+into) 来排除 Left Join 的情况
    const innerJoinRegex = /join\s+(\w+)\s+in\s+context\.(\w+)\s+on\s+(\w+)\.(\w+)\s+equals\s+(\w+)\.(\w+)(?!\s+into)/gi;

    while ((match = innerJoinRegex.exec(query)) !== null) {
        const jTable = match[2];
        const jAlias = match[1]; // Inner Join 直接用 join 后面的别名
        const leftA = match[3];
        const leftC = match[4];
        const rightA = match[5];
        const rightC = match[6];

        joinClause += `\nINNER JOIN [${jTable}] AS [${jAlias}] ON [${leftA}].[${leftC}] = [${rightA}].[${rightC}]`;
    }

    // --- 步骤 C: 解析 WHERE ---
    const whereStart = query.search(/where\s/i);
    const selectStart = query.search(/select\s/i);
    let whereClause = "";

    if (whereStart !== -1 && selectStart !== -1) {
        let whereBody = query.substring(whereStart + 6, selectStart).trim();
        
        // 处理 .Equals()
        whereBody = whereBody.replace(/(\w+)\.(\w+)\.Equals$([^)]+)$/g, convertEqualsExpression);
        // 处理 == 和 !=
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*null/g, '[$1].[$2] IS NULL');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*!=\s*null/g, '[$1].[$2] IS NOT NULL');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*==\s*(\w+)/g, '[$1].[$2] = @$3');
        whereBody = whereBody.replace(/(\w+)\.(\w+)\s*!=\s*(\w+)/g, '[$1].[$2] <> @$3');
        
        if (whereBody) whereClause = `WHERE ${whereBody}`;
    }

    // --- 步骤 D: 解析 SELECT ---
    const selectRegex = /select\s+new\s+\w+\s*$\s*$\s*\{([^}]+)\}/i;
    const selectMatch = query.match(selectRegex);
    let selectClause = "*";

    if (selectMatch) {
        const selectBody = selectMatch[1];
        const columns = [];
        const fields = selectBody.split(',');

        fields.forEach(field => {
            field = field.trim();
            if (!field) return;

            // 处理 ??
            const nullCoalesceMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)\s*\?\?\s*(.+)/);
            if (nullCoalesceMatch) {
                const colAlias = nullCoalesceMatch[1];
                const tableAlias = nullCoalesceMatch[2];
                const colName = nullCoalesceMatch[3];
                let defaultVal = nullCoalesceMatch[4].trim();
                if (defaultVal === 'string.Empty') defaultVal = "''";
                else if (defaultVal.startsWith('"') && defaultVal.endsWith('"')) defaultVal = `'${defaultVal.slice(1, -1)}'`;
                columns.push(`ISNULL([${tableAlias}].[${colName}], ${defaultVal}) AS [${colAlias}]`);
                return;
            }

            // 普通字段
            const simpleMatch = field.match(/(\w+)\s*=\s*(\w+)\.(\w+)/);
            if (simpleMatch) {
                columns.push(`[${simpleMatch[2]}].[${simpleMatch[3]}] AS [${simpleMatch[1]}]`);
            }
        });

        if (columns.length > 0) selectClause = columns.join(', ');
    }

    // --- 步骤 E: 组装 ---
    let sql = `SELECT ${selectClause}\nFROM [${mainTable}] AS [${mainAlias}]`;
    if (joinClause) sql += joinClause;
    if (whereClause) sql += `\n${whereClause};`;
    else sql += `;`;

    return sql;
}

// --- 命令行交互 ---
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
console.log('------------------------------------------------');
console.log('🚀  LINQ to SQL 混合连接修复版');
console.log('------------------------------------------------');

function startPrompt() {
    rl.question('\n👉 请输入 LINQ 语句:\n> ', (input) => {
        if (input.toLowerCase() === 'exit') { rl.close(); return; }
        try {
            const result = parseLinqToSql(input);
            console.log('\n✅ 生成的 SQL:');
            console.log('------------------------------------------------');
            console.log(result);
            console.log('------------------------------------------------');
        } catch (err) { console.log('❌ 错误:', err); }
        startPrompt();
    });
}
startPrompt();

