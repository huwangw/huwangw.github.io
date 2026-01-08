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

