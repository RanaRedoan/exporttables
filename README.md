# exporttables

**Every table of a survey dataset, in one formatted Excel workbook, one command.**

```stata
exporttables using "tables.xlsx", by(district) replace
```

`exporttables` looks at each variable, decides what kind of question it is, and
writes the right table for it:

| Variable | Table |
| --- | --- |
| Single choice (value-labelled, or 0/1) | N and % of valid answers, Total row |
| Multiple choice (SurveyCTO / ODK / Kobo `select_multiple`) | N and % of cases per option, **Valid cases (N)** row, `*` footnote |
| Continuous (other numeric) | N, Mean, Median, Mode, SD, Min, Max |
| Text with at most 50 different answers (categories typed into Excel) | N and % of valid answers, like single choice |
| Numbers stored as text (age, income from Excel) | treated as numbers: continuous, or single choice if 10 values or fewer |
| Free text (more than 50 different answers), date, identifier, form metadata, GPS, phone number, all missing | no table, listed on the Index sheet with the reason |

> [!TIP]
> **Data from Excel works too.** After `import excel`, categories such as
> `"Male"` or `"Lack of money"` are text. `exporttables` tables every text
> column with 50 or fewer different answers, and leaves out comment-style
> columns with more. Change the limit with `strmax(#)`.

With `by()`, every table becomes a cross table: answer choices down the rows,
one N / % column pair per district (or any other category), plus Total.

The workbook has two sheets:

- **Index**: dataset, date, counts of tables by type, how many text variables
  were left out, and a row per variable with a link to its table.
- **Tables**: all tables, one under another, with a navy header row, a pale blue
  Total row, a light grid, `#,##0` counts and `0.0` percentages. Every column is
  sized to fit its largest number, so nothing shows as `#####`.

Pure Stata. No Python, nothing else to install.

---

## Install

```stata
net install exporttables, ///
    from("https://raw.githubusercontent.com/RanaRedoan/exporttables/main/") replace
help exporttables
```

Requires **Stata 16 or newer**.

---

## Usage

```stata
* every variable, one-way tables
exporttables using "tables.xlsx", replace

* every variable, crossed by district
exporttables using "tables_by_district.xlsx", by(district) replace

* selected questions (a multiple-choice question is named by its parent)
exporttables gender age crops income using "selected.xlsx", by(district) replace

* a subgroup, whole-number percentages, every labelled category shown
exporttables using "female.xlsx" if gender == 2, by(district) decimals(0) allcats replace

* data kept in Excel: text columns with up to 50 answers become tables
import excel using "survey.xlsx", firstrow clear
exporttables using "tables.xlsx", by(district) replace
```

| Option | Effect |
| --- | --- |
| `by(varname)` | cross every table by this variable (numeric or string) |
| `replace` | overwrite the file |
| `sheet(name)` | name of the tables sheet (default `Tables`) |
| `decimals(#)` | decimals for percentages, 0 to 4 (default 1) |
| `allcats` | include value-label categories nobody chose (count 0) |
| `categorical(varlist)` | force these variables to be single choice |
| `continuous(varlist)` | force these variables to be continuous |
| `nomultiselect` | switch off multiple-choice detection |
| `strmax(#)` | text variables with at most # different answers get a table (default 50; `strmax(0)` for none) |

While it runs, each variable gets a line in the Results window:

```
  [1/8]      gender                        single choice    done
  [2/8]      age                           continuous       done
  [3/8]      income                        continuous       done
  [4/8]      crops (5 options)             multiple choice  done
  [5/8]      barrier                       single (text)    done
             comment                       text             skipped
             interview_date                date             skipped
  ...
  Tables exported      : 8  (single 4, multiple 2, continuous 2)
  From text variables  : 1  (categories, or numbers stored as text)
  Text variables       : 3  (no table: more than 50 different answers, or every answer different)
```

If one table fails, it is marked `FAILED` and the export carries on.

---

## Example output (`by(district)`)

**Table 5: Which crops do you grow? [crops] \***

| Response | Chattogram N | % of cases | Dhaka N | % of cases | … | Total N | % of cases |
| --- | ---: | ---: | ---: | ---: | --- | ---: | ---: |
| Rice | 52 | 52.5 | 62 | 47.3 | … | 281 | 48.0 |
| Wheat | 28 | 28.3 | 45 | 34.4 | … | 204 | 34.9 |
| Jute | 38 | 38.4 | 49 | 37.4 | … | 206 | 35.2 |
| **Valid cases (N)** | **99** | | **131** | | … | **585** | |

\* Multiple-response question: N = valid cases (respondents who answered the
question). Percentages are of valid cases and can add up to more than 100%.

**Table 3: Monthly household income (BDT) [income]**

| Statistic | Chattogram | Dhaka | … | Total |
| --- | ---: | ---: | --- | ---: |
| N | 97 | 128 | … | 575 |
| Mean | 14,411.42 | 15,603.97 | … | 14,911.32 |
| Median | 14,283.00 | 15,635.50 | … | 15,097.00 |
| SD | 5,563.05 | 5,180.87 | … | 5,039.62 |

*Not included in the statistics: -99 Don't know (n=13); -98 Refused (n=5)*

---

## How variables are classified

- **Multiple choice**: a string parent (`"1 3 5"`) plus 0/1 option variables
  `parent_code` (`crops__99` for code -99), or `q_code_k` inside a repeat
  group. Every option is checked against the parent, the same check
  [`datareport`](https://github.com/RanaRedoan/datareport) uses. If the parent
  was dropped, 0/1 variables sharing a stub (`src_1 src_2 src_3`) are grouped,
  unless they are the instances of one question in a repeat group (worded the
  same apart from a number), which keep a table each.
- **Numeric parent with dummies** (`tab fuel, gen(fuel_)`): a numeric variable
  holds one code, so `fuel` is single choice and `fuel_1 …` get no table.
- **Single choice**: numeric with a value label, or holding only 0/1 (shown as
  No / Yes).
- **Continuous**: any other numeric. A value label that names only special
  codes such as `-99 "Don't know"` does not make it categorical: those codes
  are left out of the statistics and reported under the table. Mode is left
  blank when no value occurs more than once.
- **Text**: at most 50 different answers (`strmax()`) makes a single-choice
  table, answers in natural order (`Type 2` before `Type 10`). Text whose every
  answer is a number is treated as a number. Free text with more answers,
  text in which every answer is different, dates written as text and
  `KEY` / `SubmissionDate`-style metadata get no table.
- **No table**: free text; dates (`%t` / `%d` formats, or text); identifiers (names ending
  in `id` such as `UID`, `hhid`, `resp_id`, or `key`, `uuid`, `serial`, with a
  unique value per row); form metadata (`formdef_version`, `deviceid` …); GPS
  parts (`…latitude`, `…longitude`, `…altitude`, `…accuracy`); phone numbers
  (`phone` / `mobile` / `contact` in the name, eight digits or more);
  all-missing variables.

Use `categorical()` / `continuous()` to override.

---

## Speed

The `.xlsx` is written directly as XML and zipped with Stata's `zipfile`. On a
test survey with 20,000 respondents and 94 questions crossed by 64 districts,
the export took about 8 seconds. Stata's `putexcel` creates a new cell style
for every formatted cell, which makes the same job take many minutes and can
push a workbook past Excel's 64,000-style limit.

---

## Author

Md. Redoan Hossain Bhuiyan · redoanhossain630@gmail.com
