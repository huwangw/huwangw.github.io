import com.aspose.cells.*;

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



