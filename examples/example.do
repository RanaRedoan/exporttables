* exporttables: examples
* ---------------------------------------------------------------
* A small survey-like dataset with every kind of variable.

clear
set seed 2026
set obs 500

gen district = word("Dhaka Chattogram Khulna Rajshahi Sylhet", ceil(runiform()*5))
label variable district "District"

gen byte gender = 1 + (runiform() > .5)
label define gender 1 "Male" 2 "Female"
label values gender gender
label variable gender "Gender of respondent"

gen age = 18 + floor(runiform()*50)
label variable age "Age in years"

gen income = round(rnormal(15000, 5000))
replace income = -99 in 1/15
label define income -99 "Don't know"
label values income income
label variable income "Monthly household income (BDT)"

* select_multiple, as exported by SurveyCTO: string parent + 0/1 options
gen crops = ""
foreach c in 1 2 3 4 {
    replace crops = trim(crops + " `c'") if runiform() < .4
}
replace crops = "1" if crops == ""
foreach c in 1 2 3 4 {
    gen byte crops_`c' = strpos(" " + crops + " ", " `c' ") > 0
}
label variable crops   "Which crops do you grow?"
label variable crops_1 "Rice"
label variable crops_2 "Wheat"
label variable crops_3 "Jute"
label variable crops_4 "Potato"

gen name = "Respondent " + string(_n)

* ---------------------------------------------------------------
* 1. every variable, one-way tables
exporttables using "example_tables.xlsx", replace

* 2. every variable, crossed by district
exporttables using "example_by_district.xlsx", by(district) replace

* 3. selected questions, women only
exporttables gender age crops income using "example_selected.xlsx" ///
    if gender == 2, by(district) replace
