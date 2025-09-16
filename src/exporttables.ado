

*! version 1.2.6
*! exportables.ado
*! Author: Md. Redoan Hossain Bhuiyan
*! Email: redoanhossain630@gmail.com


capture program drop exporttables
program define exporttables
    version 17
    syntax [varlist] using/

    * --- PRESERVE ORIGINAL DATA ---
    preserve

    * --- DETERMINE VARIABLES TO PROCESS ---
    if "`varlist'" == "" {
        ds
        local varlist `r(varlist)'
    }

    * --- TEMPORARILY ENCODE STRING VARIABLES ---
    local encoded_vars
    foreach var of local varlist {
        capture confirm string variable `var'
        if !_rc {
            * Create temporary encoded version
            tempvar temp_encoded
            encode `var', generate(`temp_encoded')

            * Copy value labels to original variable name for consistency
            local templabel : value label `temp_encoded'
            if "`templabel'" != "" {
                label copy `templabel' `var'
                label values `temp_encoded' `var'
            }

            * Replace original string variable with encoded numeric
            drop `var'
            rename `temp_encoded' `var'
            local encoded_vars `encoded_vars' `var'
        }
    }

    * --- SETUP EXCEL ---
    putexcel set "`using'", replace sheet("AllTables")
    local row = 1
    local tablecount = 0

    * --- LOOP OVER ALL VARIABLES TO EXPORT ---
    foreach v of local varlist {
        capture confirm variable `v'
        if _rc continue

        * --- CHECK FOR MULTI-SELECT CHILDREN ---
        ds
        local allvars `r(varlist)'
        local children ""
        foreach c of local allvars {
            if strpos("`c'", "`v'_") == 1 & regexm("`c'", ".*(_oth|_rank.*)$")==0 {
                local children `children' `c'
            }
        }

        * --- MULTI-SELECT VARIABLE ---
        if "`children'" != "" {
            * Skip if this is a child variable itself
            if regexm("`v'", ".*_[0-9]+$") | regexm("`v'", ".*_oth$") | regexm("`v'", ".*_rank.*$") {
                continue
            }

            local vlabel : variable label `v'
            if "`vlabel'" == "" local vlabel = "`v'"

            putexcel A`row' = "Variable: `v' (`vlabel')", bold
            local ++row
            putexcel A`row' = "Option", bold border(all)
            putexcel B`row' = "Frequency", bold border(all)
            putexcel C`row' = "Percent of responses", bold border(all)
            putexcel D`row' = "Percent of cases", bold border(all)
            local ++row

            * total cases = at least one child ticked
            gen byte __tmp_case = 0
            foreach c of local children {
                quietly replace __tmp_case = 1 if `c'==1
            }
            quietly count if __tmp_case==1
            local total_cases = r(N)

            drop __tmp_case

            * total responses = sum across numeric dummies
            local total_resp = 0
            foreach c of local children {
                quietly count if `c'==1
                local total_resp = `total_resp' + r(N)
            }

            * loop over children
            foreach c of local children {
                local clabel : variable label `c'
                if "`clabel'" == "" local clabel = "`c'"

                quietly count if `c'==1
                local freq = r(N)
                local pct_resp = cond(`total_resp'>0, 100*`freq'/`total_resp', .)
                local pct_cases = cond(`total_cases'>0, 100*`freq'/`total_cases', .)

                putexcel A`row' = "`clabel'", border(all)
                putexcel B`row' = `freq', border(all)
                putexcel C`row' = `=round(`pct_resp',0.01)', border(all)
                putexcel D`row' = `=round(`pct_cases',0.01)', border(all)
                local ++row
            }

            * totals row
            putexcel A`row' = "Total", bold border(all)
            putexcel B`row' = `total_resp', bold border(all)
            putexcel C`row' = 100, bold border(all)
            putexcel D`row' = "", bold border(all)
            local ++row

            * Add valid cases information only for multi-select variables
            putexcel A`row' = "Valid cases:", bold
            putexcel B`row' = `total_cases', bold
            local ++row
            local ++row
            local ++tablecount
        }
        * --- SINGLE-SELECT VARIABLE ---
        else {
            * Skip individual dummy variables of multi-select questions
            if regexm("`v'", ".*_[0-9]+$") | regexm("`v'", ".*_oth$") | regexm("`v'", ".*_rank.*$") {
                continue
            }

            local vlabel : variable label `v'
            if "`vlabel'" == "" local vlabel = "`v'"

            putexcel A`row' = "Variable: `v' (`vlabel')", bold
            local ++row
            putexcel A`row' = "Option", bold border(all)
            putexcel B`row' = "Frequency", bold border(all)
            putexcel C`row' = "Percent", bold border(all)
            local ++row

            * Get unique values and count non-missing cases for the denominator
            quietly tab `v', matrow(values)
            local total_valid = `r(N)'
            local options
            forvalues i = 1/`=r(r)' {
                local options `options' `=values[`i',1]'
            }

            local total_reported = 0
            foreach opt of local options {
                quietly count if `v' == `opt'
                local freq = r(N)
                * Only include options with a frequency greater than zero
                if `freq' > 0 {
                    local total_reported = `total_reported' + `freq'

                    local txt = "`opt'"
                    local valuelabel : value label `v'
                    * If a value label exists, use it. Otherwise, use the numeric value directly.
                    if "`valuelabel'" != "" {
                        local lbl : label (`valuelabel') `opt'
                        if "`lbl'" != "" local txt = "`lbl'"
                    }

                    local pct = cond(`total_valid'>0, 100*`freq'/`total_valid', .)
                    putexcel A`row' = "`txt'", border(all)
                    putexcel B`row' = `freq', border(all)
                    putexcel C`row' = `=round(`pct',0.01)', border(all)
                    local ++row
                }
            }

            * Total row for single-select
            putexcel A`row' = "Total", bold border(all)
            putexcel B`row' = `total_reported', bold border(all)
            putexcel C`row' = 100, bold border(all)
            local ++row
            local ++row
            local ++tablecount
        }
    }

    * --- RESTORE ORIGINAL DATA ---
    restore

    * --- Final Message ---
    di as txt "{hline 65}"
    di as txt "                     " as result "✔ EXPORT COMPLETED SUCCESSFULLY ✔"
    di as txt "{hline 65}"
    di as txt "   Number of tables created : " as result `tablecount'
    di as txt "   File saved as            : " as result "`using'"
    di as txt "{hline 65}"
    di as txt "     Thank you for using " as result "exporttables" as txt "!"
    di as txt "{hline 65}"
end
