{smcl}
{* *! version 2.2.0  10oct2026}{...}
{viewerjumpto "Syntax" "exporttables##syntax"}{...}
{viewerjumpto "Description" "exporttables##description"}{...}
{viewerjumpto "Options" "exporttables##options"}{...}
{viewerjumpto "How variables are treated" "exporttables##rules"}{...}
{viewerjumpto "Examples" "exporttables##examples"}{...}
{viewerjumpto "Stored results" "exporttables##results"}{...}
{title:Title}

{phang}
{bf:exporttables} {hline 2} Export a table for every variable of a survey
dataset to one formatted Excel workbook, one-way or crossed by district


{marker syntax}{...}
{title:Syntax}

{p 8 17 2}
{cmd:exporttables}
[{varlist}]
{ifin}
{cmd:using} {it:filename}
[{cmd:,} {it:options}]

{synoptset 24 tabbed}{...}
{synopthdr}
{synoptline}
{synopt:{opt by(varname)}}cross every table by {it:varname}: one column block per category (for example each district) plus Total{p_end}
{synopt:{opt replace}}overwrite {it:filename} if it exists{p_end}
{synopt:{opt sheet(name)}}name of the sheet holding the tables; default {cmd:Tables}{p_end}
{synopt:{opt dec:imals(#)}}decimals shown for percentages, 0 to 4; default {cmd:1}{p_end}
{synopt:{opt all:cats}}show every category defined in the value label, even if nobody chose it{p_end}
{synopt:{opt cat:egorical(varlist)}}treat these variables as single choice{p_end}
{synopt:{opt cont:inuous(varlist)}}treat these variables as continuous{p_end}
{synopt:{opt nomulti:select}}do not detect multiple-choice questions{p_end}
{synopt:{opt str:max(#)}}text variables with at most {it:#} different answers get a table; default {cmd:50}{p_end}
{synoptline}


{marker description}{...}
{title:Description}

{pstd}
{cmd:exporttables} writes one table per question into a single Excel sheet
and adds an {cmd:Index} sheet that lists every variable, what was done with
it, and a link to its table.  With no {varlist}, every variable in the
dataset is processed.

{pstd}
The kind of table follows the kind of variable:

{p2colset 8 30 32 2}{...}
{p2col:{it:single choice}}N and % of valid answers, with a Total row{p_end}
{p2col:{it:multiple choice}}N and % of cases for each option, a
{bf:Valid cases (N)} row, and a {bf:*} footnote explaining that the
percentages can add up to more than 100{p_end}
{p2col:{it:continuous}}N, Mean, Median, Mode, SD, Min and Max{p_end}
{p2col:{it:text}}a single-choice table when it has at most {opt strmax()}
(50) different answers, such as categories typed into Excel; no table for
free text with more answers than that{p_end}
{p2colreset}{...}

{pstd}
With {opt by()}, the same tables are produced as cross tables: the rows are
the answer choices (or statistics) and the columns are the categories of the
{opt by()} variable plus a Total column.  Percentages are column
percentages.  Observations missing on the {opt by()} variable are left out
of every table, and a note says how many.

{pstd}
While it runs, {cmd:exporttables} prints one line per variable showing the
kind of table and whether it was {bf:done} or {bf:skipped}.  A table that
fails is reported as {bf:FAILED} and the export carries on with the next
one.

{pstd}
Tables have a navy header row, a pale blue Total row and a light grid, and
every column is made wide enough for its largest number.  The workbook is
written directly as Office Open XML and zipped with {help zipfile}.  No
Python and no other software is needed, and large exports stay fast because
every cell shares a small set of styles.


{marker options}{...}
{title:Options}

{phang}
{opt by(varname)} names the column variable, numeric or string.  Column
headings come from its value label, or from its values.  At most 1,000
categories.

{phang}
{opt replace} overwrites an existing file.  Close the file in Excel first.

{phang}
{opt sheet(name)} renames the tables sheet.  {cmd:Index} is reserved.

{phang}
{opt decimals(#)} sets the decimals shown for percentages.  Percentages are
also rounded to this many decimals, so copied values match what is shown.

{phang}
{opt allcats} adds the categories defined in a value label but never
chosen, with a count of 0, so that tables keep the same shape across
datasets.

{phang}
{opt categorical(varlist)} and {opt continuous(varlist)} override the
automatic choice described below.

{phang}
{opt nomultiselect} turns off multiple-choice detection; option dummies are
then tabulated one by one as yes/no questions.

{phang}
{opt strmax(#)} sets how many different answers a text variable may have and
still get a table; default 50.  {cmd:strmax(0)} gives no tables for text.  A
text variable named in {opt categorical()} always gets a table.


{marker rules}{...}
{title:How variables are treated}

{pstd}
{bf:Multiple choice.}  A select_multiple question from SurveyCTO, ODK or
Kobo arrives as a string {it:parent} holding the codes chosen ({cmd:"1 3 5"})
and one 0/1 variable per option, named {it:parent}{cmd:_}{it:code}
({cmd:crops_1}, {cmd:crops_3}, ...; code -99 is {cmd:crops__99}).  Inside a
repeat group the names are {it:q}{cmd:_}{it:code}{cmd:_}{it:k}.  Each option
variable is checked against the parent before the question is accepted, the
same check {cmd:datareport} uses.  If the parent was dropped, option
variables sharing a stub ({cmd:src_1 src_2 src_3}) and holding only 0 and 1
are grouped, unless they look like the instances of one question asked in a
repeat group ({cmd:loan_1 loan_2 loan_3} worded the same apart from a number,
or sharing a value label with codes other than 0 and 1); those keep a table
each.  Valid cases are the respondents who answered: parent not missing, or,
with no parent, not missing on at least one option.

{pstd}
Option labels come from the option variable's label.  When every option
label starts with the same {cmd:"Question: "} text, that text becomes the
table title and is removed from the options.

{pstd}
{bf:Numeric parents.}  A numeric variable can hold only one code, so
{cmd:fuel} with {cmd:fuel_1 fuel_2 fuel_3} (as made by
{cmd:tab fuel, gen(fuel_)}) is a single choice: {cmd:fuel} gets the table and
the indicators are listed on the Index sheet without a table.

{pstd}
{bf:Single choice.}  Numeric variables with a value label, and numeric
variables holding only 0 and 1 (shown as No / Yes).

{pstd}
{bf:Text.}  Data typed into Excel, or imported with {cmd:import excel},
often keeps its categories as text ({cmd:"Male"}, {cmd:"Lack of money"}).
A text variable with at most {opt strmax()} different answers (50 by
default) gets a single-choice table, its answers in natural order
({cmd:Type 2} before {cmd:Type 10}).  A text variable whose every answer is
a number is treated as a number: with more than 10 different values (age,
income) it is continuous, with 10 or fewer (a 1 to 5 rating) it is single
choice.  No table is made for text with more than {opt strmax()} different
answers (comments, names, addresses), for text in which every answer is
different, for dates and times written as text, and for form metadata such
as {cmd:KEY} or {cmd:SubmissionDate}.  On the Index sheet these tables are
marked {bf:(text)}.

{pstd}
{bf:Continuous.}  Other numeric variables.  A value label that names only a
few special codes ({cmd:-99 "Don't know"}) does not make a variable
categorical: a labelled variable with more than 10 unlabelled values is
continuous, the labelled codes are left out of the statistics, and a note
under the table lists them with their counts.  Mode is the most frequent
value (the smallest one when several tie); it is left blank when no value
occurs more than once.

{pstd}
{bf:No table.}  Free text (see {bf:Text} above); dates and times (a
{cmd:%t} or {cmd:%d} format, or written as text); identifiers (a name ending in {cmd:id} such as {cmd:UID},
{cmd:hhid} or {cmd:resp_id}, or named {cmd:key}, {cmd:uuid} or
{cmd:serial}, with a different value in every row); form metadata
({cmd:formdef_version}, {cmd:deviceid}, {cmd:simid} ...); the parts of a GPS
reading (names ending in {cmd:latitude}, {cmd:longitude}, {cmd:altitude} or
{cmd:accuracy}); phone numbers (a name containing {cmd:phone},
{cmd:mobile}, {cmd:contact} or {cmd:cell}, with values of eight digits or
more); 0/1 indicators of a numeric parent; variables with every value
missing.


{marker examples}{...}
{title:Examples}

{pstd}All variables, one-way tables{p_end}
{phang2}{cmd:. exporttables using "tables.xlsx", replace}{p_end}

{pstd}All variables, crossed by district{p_end}
{phang2}{cmd:. exporttables using "tables_by_district.xlsx", by(district) replace}{p_end}

{pstd}Selected questions only; a multiple-choice question is named by its parent{p_end}
{phang2}{cmd:. exporttables gender age crops income using "selected.xlsx", by(district) replace}{p_end}

{pstd}Female respondents, whole percentages, every label category shown{p_end}
{phang2}{cmd:. exporttables using "female.xlsx" if gender == 2, by(district) decimals(0) allcats replace}{p_end}

{pstd}Data kept in Excel, where every column comes in as text{p_end}
{phang2}{cmd:. import excel using "survey.xlsx", firstrow clear}{p_end}
{phang2}{cmd:. exporttables using "tables.xlsx", by(district) replace}{p_end}

{pstd}The same, allowing up to 100 different answers per text column{p_end}
{phang2}{cmd:. exporttables using "tables.xlsx", by(district) strmax(100) replace}{p_end}


{marker results}{...}
{title:Stored results}

{pstd}
{cmd:exporttables} stores the following in {cmd:r()}:

{synoptset 20 tabbed}{...}
{p2col 5 20 24 2: Scalars}{p_end}
{synopt:{cmd:r(N)}}observations used{p_end}
{synopt:{cmd:r(N_tables)}}tables exported{p_end}
{synopt:{cmd:r(N_single)}}single-choice tables{p_end}
{synopt:{cmd:r(N_multiple)}}multiple-choice tables{p_end}
{synopt:{cmd:r(N_cont)}}continuous tables{p_end}
{synopt:{cmd:r(N_text)}}tables made from text variables{p_end}
{synopt:{cmd:r(N_string)}}text variables left out (too many different answers){p_end}
{synopt:{cmd:r(N_skipped)}}other variables left out{p_end}
{synopt:{cmd:r(N_failed)}}tables that failed{p_end}

{p2col 5 20 24 2: Macros}{p_end}
{synopt:{cmd:r(using)}}file written{p_end}
{synopt:{cmd:r(by)}}the {opt by()} variable{p_end}


{title:Requirements}

{pstd}
Stata 16 or newer.


{title:Author}

{pstd}
Md. Redoan Hossain Bhuiyan{break}
redoanhossain630@gmail.com{break}
{browse "https://github.com/RanaRedoan/exporttables"}


{title:Also see}

{psee}
{help tabulate oneway}, {help tabstat}, {help putexcel}, {help zipfile}
{p_end}
