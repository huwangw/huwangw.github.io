3import com.aspose.cells.*;

public void setupColumnDataValidations(Worksheet worksheet, int startRow, int endRow) {
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
List<DataValidationUtil.ValidationInfo> validations = DataValidationUtil.getValidationsInRange(worksheet, targetRange);

// 3. 清除旧的（解决报错的关键）
worksheet.getValidations().clear(10, 0, 4991, 3); // 清除目标区域旧数据

// 4. 重新创建
for (DataValidationUtil.ValidationInfo info : validations) {
    DataValidation newDV = worksheet.getValidations().add();
    
    // 复用原有的公式 (如 "=NamedRange")
    newDV.setFormula1(info.validation.getFormula1());
    newDV.setFormula2(info.validation.getFormula2());
    newDV.setType(info.validation.getType());
    // ... 复制其他属性 ...

    // 设置新的区域（注意：这里可以只设置该列，或者设置具体的交集）
    CellArea newArea = new CellArea();
    newArea.StartRow = 10;
    newArea.EndRow = 5000;
    newArea.StartColumn = info.effectiveArea.StartColumn; // 获取到的列
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
            // 一个验证规则可能应用在不连续的多个区域上
            CellArea[] areas = validation.getAreas();
            if (areas == null) continue;

            for (CellArea area : areas) {
                // 遍历该区域内的每一列
                // 这样可以确保如果一个规则覆盖了 A1:C10，我们能知道 A、B、C 三列都有这个规则
                for (int colIdx = area.StartColumn; colIdx <= area.EndColumn; colIdx++) {
                    ColumnValidationRule rule = new ColumnValidationRule();
                    rule.columnIndex = colIdx;
                    rule.columnName = worksheet.getCells().getColumnName(colIdx); // 将索引转为 A, B, AA 等
                    rule.validationType = validation.getType().toString();
                    rule.operator = validation.getOperator().toString();

                    // 获取公式1 (通常是数据源或最小值)
                    Object formulaObj1 = validation.getFormula1();
                    rule.formula1 = (formulaObj1 != null) ? formulaObj1.toString() : "";

                    // 获取公式2 (如果是 "介于" 或 "不介于" 等操作符)
                    Object formulaObj2 = validation.getFormula2();
                    rule.formula2 = (formulaObj2 != null) ? formulaObj2.toString() : "";

                    result.add(rule);
                }
            }
        }

        // 可选：根据列索引去重，防止同一列被多个区域重复添加
        // 如果你需要保留原始的区域信息则不需要去重
        return removeDuplicates(result);
    }

    // 简单的去重方法，确保同一列只显示一次（取第一个遇到的规则）
    private static List<ColumnValidationRule> removeDuplicates(List<ColumnValidationRule> list) {
        Map<Integer, ColumnValidationRule> map = new LinkedHashMap<>();
        for (ColumnValidationRule rule : list) {
            // 如果该列还没有记录，则放入 map
            if (!map.containsKey(rule.columnIndex)) {
                map.put(rule.columnIndex, rule);
            }
        }
        return new ArrayList<>(map.values());
    }
}







// 加载模板
Workbook workbook = new Workbook("template.xlsx");
Worksheet templateSheet = workbook.getWorksheets().get(0); // 获取第一个Sheet

// 调用分析方法
List<TemplateValidationAnalyzer.ColumnValidationRule> rules = TemplateValidationAnalyzer.analyzeValidationsInSheet(templateSheet);

// 打印结果
System.out.println("模板中使用了数据验证的列：");
for (TemplateValidationAnalyzer.ColumnValidationRule rule : rules) {
    System.out.println(rule.toString());
}





