# Demo Flexcel Report
Demo use Flexcel Report to export repport in Asp.NET project. Use pdfjs library to view.
> *Copyright 2021 [ChungNA](https://github.com/chungpj)*
> 
> ## Getting started
> 1. Clone this repository:
```
---
git clone https://github.com/chungpj/demo.flexcelReport.git
---
```
2. Run 
```
---
Rebuild and Run project Demo.UI
---
```
3. Docs
```
---
1. Create template file (word, excel..)
  1.1. config Dynamic Params, TableDatas
  1.2. config TableDatas RelationShip (if has)
  1.3. config Formulas (TableDatas Range)
2. Create Report and ReportFilter class (mapping Params, TableDatas from code and template file)
3. Create Controller and Action return file content (byte[])
4. Call Action
---
```

## Config Demo:
![config_table_data](/config_table_data.png)

![config_relationship](/config_relationship.png)

![config_formulas](/config_formulas.png)

![code_config_table_data](/code_config_table_data.png)
## See Demo Result:
![parent_data](/parent.png)

![childs_data](/childs.png)

![demo](/demo.png)

![print](/print.png)
