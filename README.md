
## Table of Contents
- [Table of Contents](#table-of-contents)
- [Overview](#overview)
- [Application](#application)
- [Workflow](#workflow)
  - [Example Workflow: Concrete, forming.](#example-workflow-concrete-forming)
- [Interface \& Functional Components](#interface--functional-components)
  - [Required fields:](#required-fields)
  - [Buttons:](#buttons)
  - [Sheets:](#sheets)
  - [Advanced Features](#advanced-features)
    - [Statistical Outputs](#statistical-outputs)
    - [Relationship based Scheduling](#relationship-based-scheduling)
    - [Distribution Types and Descriptions](#distribution-types-and-descriptions)
    - [Calendars](#calendars)
    - [Limitations](#limitations)
    - [Code Explaination](#code-explaination)
- [Author](#author)
     

## Overview

The Resource Loader Tool enables users to distribute resources across a defined time period using customizable distribution models and calendar settings. It supports a wide range of planning and forecasting needs, including construction activities, labor planning, equipment usage, and financial analysis.

## Application
This tool can be used for:

- Steel tonnage installation  
- Concrete cubic yardage placement  
- Square foot concrete forming  
- Cash flow analysis  
- Equipment usage tracking  
- Manhour distribution across trades  

## Workflow

1. Input activities, and descriptions
2. Provide a schedule
3. Input resource
4. Select distribution type
5. Select Calendar
6. Click Schedule
   
### Example Workflow: Concrete, forming.
1. Input concrete forming activities.
    - Must have unique activity ID
    - Description headers could be: Job#, Building, Level, Subcontractor, Category, Slab Number, unit of measure etc. These should correspond to required BI input.*
2. Provide a schedule for forming the pours.
    - Or generate a schedule by clicking the generate schedule box below the "Schedule" button, then input durations and the first start date.
3. Apply the resource.
    - Forming square foot per pour area.
4. Select the distribution.
    - Must be one of the drop down* values in the distrubution column drop down list. 
    - Linear, is the most intuitive. The distribtuion types can be viewed on the "Distributions" tab
5. Choose the calendar.
    - Must be one of the drop down* that corresponds with the schedule. 
6. Click the "Schedule" button.
      - Once data is input click the "Schedule" button and the resource per day will be calculated and populated on the gantt for each of the activities. 
      - This data is automatically prepared for export to a Power BI dashboard on the "Export" tab, and is also available for visualization.
      - Clicking the "Visualize" button will open the "Dashboard" tab showing your plot of resource distribution over time. You can choose to show data lables or a trend line as well.

*The schedule values can be copied from P6, another workbook, manually input, or calculated. However, drop down fields such as "Distribution" and "Calendar", are data validated. Your selection must be one from the drop down list. You can copy and paste, or drag down in those fields these values as well, but they are limited to the values on the list.*

## Interface & Functional Components
### Required fields:
 - Activity ID
 - Description Fields*
 - Duration
 - Start
 - Resource
 - Distribution
 - Calendar

*The column headers between "Actvity ID" and "Duration", or Description fields are flexible and user-defined. They appear in the Export tab and should match Power BI import headers.*

*Non-Required fields include any of the columns between Activity ID and Duration as well as the Finish field. The finish date will be overwritten to reflect the finish date calculated from the start date using the selected calendar.*

### Buttons:
 - Schedule
   - Calculate Relationships*
 - Visualize
   - Show Data Lables*
   - Show Trend Line*
 - Reset Data

*See Advanced Features below*

The "Schedule" button and corresponding check box will execute the core functions. taking the data and project it out based on distribution and calendar selections. 

Once clicking on the schedule button and the calculations are complete, the fields for MAX, MIN, the dates the occur on, AVE, SLOPE (linear interpolation), and the R Squared (RSq) Correlation ratio (0-1) will be diplayed for the current scheduled data. 

The "Visualize" button opens up the "Dashboard" tab, showing a chart containing the cumulative distribution of that Resource over time. Checkboxes are available to show data labels, and trend line on the chart.

The "Reset" button clears the **ALL** the fields and current plots.

### Sheets:
 - Gantt
 - Distributions
 - Export
 - Dashboard
 - (Hidden)
   - Calendars
   - Holidays
   - Pivot
   - Plot
   - ProductionChart
   - Distribution Tables

The "Gantt" sheet is the user will input data, and where the data will be projected.

"Distributions" is a sheet containing a table that provides a visualization of the various distribution types.

"Export" contains the calculated data in a format for use with BI dashboards.

"Dashboard" is a sheet containing a chart of plotted data.

Sheets have been hidden to avoid unknowinly disrupting the program. 
- Calendars contains a few tables where the calendars are represented, ie. 5 day, 6 day, 7 day no holiday. It also stores the graphical overlay, to be restored if deleted.
- Holidays contains a table with a list of all the typical holidays up to New Years Eve 2030.
- Pivot contains the pivot table connected to the pivot chart.
- Plot containts a table with the dates and running total of the daily resource.
- ProductionChart contanes a chart that grapically represents the data on Plot.
- Distribution Tables contains a table with all the distribution charts plotted from 1 to 1000. The charts are normalized values, so the sum of each plot will be equal to 1.


### Advanced Features
#### Statistical Outputs
Key data points such as MAX, MIN, AVE, values along with SLOPE (linear interpolation) an approximate value of resource required per day in order to achieve the projection, and the R Squared (RSq) a value (0-1, 0 not corelated) that tells you how dependable the SLOPE value is. 

#### Relationship based Scheduling
You can use the built in scheduling tools to generate the schedule data based on an initial start date, durations, and relationship information, such as Finish Start (FS) or Start Start (SS) relationships with Lag. 
The check box below the schedule button will enable schedule calcuation based on relationships, selecting the box will enable 3 addtional columns:
 - Rel. *(Relationship Type. Default: FS)*
 - w/ *(What activity is the relationship with. Default: Previous Activity, Prev)*
 - Lag *(The lag time befor the start of the related activity. Default: 0)*

#### Distribution Types and Descriptions
 - Linear
   - Resource is divided evenly by the duration.
 - Bell Curve
   - Resource utilization ramps up sharply, peaks in the middle of the duration then down.
 - Double Peak
   - Resource utilization ramps up sharply then down, experiences a local minimum at mid duration then peaks again with another sharp decline.
 - Back Loaded
   - Resource utilization is weighed higher in last of the duration.
 - Front Loaded
   - Resource utilization is weighed higher in the first half duration.
 - Trapazoidal
   - Resource utilization ramps up then settles into linear utilization for about half the overall duation then back down.
 - Smooth Trap
   - Resource utilization begins immedatly with a slow, slight increase in utilization then slowly decreases slightly lower than initial utilization.
 - Steel Fab
   - Resource utilization begins late, however experiences a sharp and steady increase in utiliztaion to the end of the duration. Continual increase.
 - Soft Front
   - Similar to Front Loaded, with the difference being a more gradual decrease in utilization over the duration. Near to linear.

#### Calendars
The calendars included are:
There are 6 calendars included, and room for custom calendars.
 - 5 day work week with and without holidays.
 - 6 day work week with and without holidays.
 - 7 day work week with and without holidays.

Holidays are standard, and more can be added if necessary.

#### Limitations
Activity duration should not exceed 1000 days (days or any other unit of duration).
If necessay to exceed the duration limiation, break the activitiy into smaller components.

User added columns may only be added beween the *Activity ID* and *Duration* columns. The column order of the user added columns will be the same order for the export.

MAX, and MIN, will take the earliest occuurance IF there are multiple days with those values. SLOPE, and RSq only reflect linear correlation. If RSq is a low corelaction (Closer to 0) then understand the the linear representation may not be suitable for your purposes.


#### Code Explaination
```
**GanttModule**
Core Functions:
    SumByDistributions
        -This function takes the distrubtion plot slection along with the duration and returns an array of len(duration) with each arr(i) having an equal distribution of points.
    CreateDistributionArray
        -This function multiplies the resource value and the SumByDistributions array, creating a new array equal to the distribution of resouses per duration.
    AssignDistributionToDates
        -This function takes the CreateDistributionArray and accounts for workdays and holidays.
    CalcuateStartAndFinish
        -This function takes the start date and duration and caculates the end date, then uses the relationship data (if any) to calculate the subsequent activity start date.
    GetDateCustom
        -This function parces through the duration using calendar workdays and holidays, returns a date from its calcuation. This will be used in the schedule calculation.
    PopulateDateHeaders
        -This populates and ensures proper format for the dates along the gantt header. It looks for the earliest start date and latest finish date as its range.
    WriteToCells
        - This function writes the values of CreateDistributionArray to the gantt table
    RunResources
        -This is the subroutine that executes the functions in order.

    Helper Functions:
        InitializeColumnIndex
            -This inializies all the columns at their current locations.
        BuildColumnIndex
            -Generates an index of columns. If the user added columns this will allow the referenced columns be accessed.
        IDFinishDate
            -This function will use the schedule logic to calcuate the finish date if using an acitivy ID in the w/ column, used in in tandum with GetDateCustom.
        IDStartDate
            -This function will use the schedule logic to calcuate the Start date if using an acitivy ID in the w/ column, used in in tandum with GetDateCustom.
        ClearAllInteriorColor
            -Clears the form cell colors
        ClearFromJ2
            -Clears the form
            -This function returns a column of data based on table selection
        DeleteEmptyRows
            -Loops through the target WS looking for empty rows, deletes the empty rows.
        ColumnLetter
        ActivateGanttSheet
        ShowMyForm 
        GetLastColumnRange 
        DebugDateRange 
        DiagnoseClearContentsIssue 
        FineDateColumn 
        FormatDateRange 
        GetTableRange 
        GetColumn 

**Dashboard Module**
Core Functions:
    CreatePivotChartWithSlicerPanel
        -This function creates a pivot chart with slicers from the activity description columns
    AddLabelsAndTrendlineToPivotChart
        -This function evaluates the state of the checkboxes below the "Visualize" button and will apply data lables or trendline respectively

    Helper Functions:
        PivotChartSytleChecker
            -Prints out the current selected plot style number, used to hardcode that style in the CreatePivotChartWithSlicerPanel function.


**Plotting Module** (No longer relevent with Dashboard)
Core Functions:
    SumValuesByDate
        -This function takes the sum of all the calculated values by date, it stores the data on the plot sheet
    CreateChartOnNewSheet
        -This function creates a new plot from the data generated from SumValuesByDate on the plot sheet.
    GetTrandlineStats
        -Calculates and populates the data for slope and RSq
    RunChartting
        -This subroutine executes the functions in order
    
    Helper Functions:
        UpdateChartFormatting 
            - Checks for the "data lables" and "trend line" check boxes and adds them to the plot based on their status.
        FormatChart 
            -Formats the chart
        GetLastUsedColumnPair 
        SheetsExists 
```

## Author
Steve Hall P.E.  
shall@austin-ind.com