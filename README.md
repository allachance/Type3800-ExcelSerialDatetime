# Type3800

Custom TRNSYS component that converts simulation time into the Excel Serial Datetime format.

## Purpose

Type 3800 bridges TRNSYS simulation outputs with spreadsheet tools such as Microsoft Excel by converting TRNSYS simulation time into Excel’s serial datetime format. This makes it easier to analyze, graph, and post-process TRNSYS results directly in Excel.

## Requirement

- TRNSYS v.18

## Installation

1. Clone the repo or download the source files:

   ```bash
   git clone https://github.com/allachance/TRNSYS-ExcelSerialDatetime-Type3800.git
   ```

2. Copy the folders into your TRNSYS 18 installation directory, e.g., ```C:\TRNSYS18```

## Configuration

### Parameters

| Name | Description | Options |
|-|-|-|
| Mode        | Defines how multi-year simulations are handled | `1` → Time accumulates across years<br>`2` → Time loops within the reference year |
| Reference Year | Simulation starting year               | —                                                                       |

### Output
| Name          | Description                             | Options |
|-|-|-|
| Serial Date | Serial number compatible with Excel datetime | —       |


> [!NOTE]  
> TRNSYS simulations always assume a non-leap year with 8,760 hours per year. Therefore, February 29th is never simulated and will be skipped during conversion. To ensure a continuous date mapping to Excel serial numbers, it is recommended to avoid using leap years as the reference year.

## Usage

Use the following formula to convert the ```Serial Date``` number into a datetime format:

```excel
=TEXT(A1,"yyyy-mm-dd hh:mm:ss")
```
