**{smcl}
*{hline}
* {bf:exporttables} - Export single-select and multi-select survey tables to Excel
*{hline}

{title:Description}

{p 4 4 2}
{bf:exporttables} is a Stata program designed to automatically export summary tables for all variables (or a selected subset) in a survey dataset to a Microsoft Excel file (.xlsx). It intelligently detects and handles both {it:single-select} and {it:multi-select} questions, calculating appropriate frequencies and percentages for each type. The output is formatted for easy interpretation and reporting.

{p 4 4 2}
For multi-select questions (represented as a set of dummy variables), it calculates both the percentage of responses and the percentage of cases, providing a complete picture of the results.

{title:Syntax}

{p 8 15 2}
{cmd:exporttables}
[{varlist}]
{cmd:,} {bf:using}({it:filename}) 

{title:Options}

{synoptset 20 tabbed}{...}
{synopt:{opt using}({it:string})}specifies the path and filename for the output Excel file. The {opt .xlsx} extension is recommended. This option is {bf:required}.{p_end}

{title:Remarks}

{p 4 4 2}
The command performs the following operations automatically:

{p 6 6 2}
1.  {bf:Variable Selection:} If no {varlist} is specified, the command processes all variables in the dataset.{p_end}
{p 6 6 2}
2.  {bf:Multi-Select Detection:} For a given variable, it checks for the existence of other variables whose names start with {it:varname_}. These are treated as the dichotomous (0/1) choice variables for a multi-select question. Variables ending in {bf:_oth} or {bf:_rank} are automatically excluded from this check.{p_end}
{p 6 6 2}
3.  {bf:Table Creation:}
{p_end}
{p 9 9 2}
-   {bf:Multi-Select Tables:} Include columns for Frequency, % of Responses, and % of Cases.{p_end}
{p 9 9 2}
-   {bf:Single-Select Tables:} Include columns for Frequency and Percent. Variables must have a value label attached to be processed.{p_end}
{p 6 6 2}
4.  {bf:Export:} All tables are written to a single sheet named "AllTables" in the specified Excel file, with clear titles, bold headers, and borders for readability.{p_end}


{title:Examples}

{p 4 4 2}
Export all labeled single-select and multi-select variables to an Excel file.

{phang2}{cmd:. exporttables using "My_Survey_Tables.xlsx"}{p_end}

{p 4 4 2}
Export only the variables 'gender', 'age_group', and the multi-select question whose root name is 'social_media'.

{phang2}{cmd:. exporttables gender age_group social_media using "Selected_Tables.xlsx"}{p_end}


{title:Notes on Data Structure}

{p 4 4 2}
{bf:Multi-Select Questions:} Must be coded as sets of dummy variables where 1 = selected and 0 = not selected. The variable names must follow the pattern {it:rootname_optionname}. For example, a question "Which platforms do you use?" with root name {bf:platforms} should have variables like:
{break}    {bf:platforms_facebook}
{break}    {bf:platforms_twitter}
{break}    {bf:platforms_instagram}

{p 4 4 2}
{bf:Single-Select Questions:} Must be numeric variables with value labels attached. The command will skip numeric variables without value labels and string variables.

{title:Author}

{p 4 4 2}
{bf:Author:} Md. Redoan Hossain Bhuiyan{p_end}
{p 4 4 2}
{bf:Email:} redoanhossain630@gmail.com{p_end}

{title:Also See}

{p 4 4 2}
Help for {help putexcel}, {help ds}, {help levelsof}{p_end}

*{hline}
**