# IT Expenditure Variance Analysis

An Excel analysis of company IT expenditure against plan and forecast. The workbook adds a **Variance Analysis** sheet to the source workbook, which also contains the `Data` and `Future Data` sheets.

## Analysis

The analysis compares expenditure by month and across four dimensions:

- Business Area
- Country
- IT Area
- Cost element name

It reports Actual, Forecast, Plan, Actual vs Plan, Actual variance percentage, and Forecast vs Plan.

Actuals are combined from `Data` and `Future Data`. Forecast and Plan are taken from `Data`. The summary uses `SUMIFS` to aggregate each measure by month or category before calculating variance, because Actual, Forecast, and Plan are stored in separate sparse rows.

## Key result

Full-year Actual is **866.4 million source units**, compared with Plan of **900.4 million**. Actual expenditure is **34.0 million (3.8%) below plan**.

The workbook also compares the Forecast with Plan by month and category. A variance close to zero overall can conceal larger over- and underspends within individual cost elements or business areas.

## Formula logic

**Actual variance**

```excel
Actual - Plan
```

A positive result means Actual is greater than Plan. A negative result means Actual is less than Plan.

**Actual variance percentage**

```excel
=IF(Plan=0,"n.m.",Variance/ABS(Plan))
```

`n.m.` means the percentage is not meaningful because Plan is zero.

**Monthly Actual example**

```excel
=SUMIFS(Data!$J:$J,Data!$A:$A,$A2)
 +SUMIFS('Future Data'!$J:$J,'Future Data'!$A:$A,$A2)
```

**Monthly Plan example**

```excel
=SUMIFS(Data!$L:$L,Data!$A:$A,$A2)
```

In the source sheets, column A contains the month, J contains Actual, K contains Forecast, and L contains Plan on `Data`. `Future Data` contains additional Actuals.

## How to use

Open the workbook in Excel and go to **Variance Analysis**. Review the monthly comparison first, then use the dimension tables to identify categories contributing to over- or underspend. Positive variance values are above plan; negative values are below plan.

## Limitations

- The source does not specify the reporting year, currency, or unit scale. Amounts are therefore shown in source units.
- The source identifies where variances occur, but does not include the commitments, vendors, project changes, or explanations needed to determine why they occurred.
- Negative amounts may represent credits or reversals. Confirm the source sign convention before interpreting them as savings or overspend.
- The variance percentage is unavailable when Plan is zero.
