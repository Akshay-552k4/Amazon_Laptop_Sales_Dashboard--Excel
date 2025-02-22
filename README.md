# Amazon_Laptop_Sales_Dashboard--Excel
This dashboard analyzes Amazon's laptop sales. Data was cleaned and transformed for insights. Created Pivot tables, multiple charts, slicers, and automation to enhance interactivity and analysis.
# **Amazon Laptop Sales Report - Excel Dashboard**  
    Application.Calculation = xlCalculationManual
    
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Pivot Tables Updated Successfully!", vbInformation, "Refresh Complete"
End Sub
```

### **3️⃣ Refresh Button Macro for Manual Updates**  
```vba
Sub RefreshAllPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Dashboard Updated Successfully!", vbInformation, "Refresh Complete"
End Sub
```

### **4️⃣ Auto-Sort Pivot Table by Sales**  
```vba
Sub AutoSortPivotTable()
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set ws = ThisWorkbook.Sheets("YourPivotSheetName") ' Change sheet name
    Set pt = ws.PivotTables("YourPivotTableName") ' Change pivot table name
    
    pt.PivotFields("Sales").AutoSort xlDescending, "Sum of Sales"
    
    MsgBox "Pivot Table Sorted Successfully!", vbInformation, "Sorting Complete"
End Sub
```

---

## 📊 **Final Output: Interactive Excel Dashboard**  
The **Amazon Laptop Sales Dashboard** provides a consolidated view of sales, brand performance, and stock availability. It includes **automated pivot table updates, real-time sorting, and an interactive navigation experience** to make data-driven decisions efficiently.  

📊 **Dashboard Preview:**  
_(Attach an image or GIF of the Excel dashboard here)_  

---

## 🔗 **Conclusion**  
This **Excel-based Amazon Laptop Sales Dashboard** is designed for easy navigation, insightful analysis, and automation. By leveraging **Pivot Tables, Charts, and VBA Macros**, the dashboard delivers actionable insights into sales performance and inventory management.  

✅ **Data-Driven Decision Making**  
✅ **Automated Analysis & Sorting**  
✅ **User-Friendly Interactive Visuals**  

---

## 📜 **License**  
This project is licensed under the **MIT License**. Feel free to modify and use it for your analytical needs.  
