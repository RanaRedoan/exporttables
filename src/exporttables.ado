*! version 2.2.0  10oct2026
*! exporttables: export a frequency / summary table for every variable
*!               of a dataset to one formatted Excel sheet
*! Author: Md. Redoan Hossain Bhuiyan
*! Email : redoanhossain630@gmail.com
*
* What it builds, one table per question:
*
*   single choice    (value-labelled or 0/1 numeric,  -> N and % of valid
*                     or text with at most strmax()
*                     different answers)
*   multiple choice  (select_multiple option dummies) -> N and % of cases,
*                                                        valid cases (N) row
*                                                        and a * footnote
*   continuous       (any other numeric, or numbers   -> N, Mean, Median,
*                     stored as text)                    Mode, SD, Min, Max
*   free text, date, ID, metadata, all-missing        -> no table, listed on
*                                                        the Index sheet
*
* With by(), every table becomes a cross table: one column block per
* category of the by() variable (for example each district) plus Total.
*
* Multiple-select detection follows -datareport-: a select_multiple
* exports as a parent (codes picked, e.g. "1 3 5") plus one 0/1 dummy per
* option, either P_<code> or, inside a repeat group, Q_<code>_<k>.  The
* naming scheme is chosen once per question and every dummy is checked
* against the parent before the block is accepted.  Dummies whose parent
* was dropped are grouped by their common stub instead.

program define exporttables, rclass
    version 16

    syntax [varlist(default=none)] [if] [in] using/ [,   ///
        BY(varname) SHEET(string) REPLACE ALLcats          ///
        DECimals(integer 1) CATegorical(varlist)           ///
        CONTinuous(varlist) noMULTiselect STRmax(integer 50) ]

    *------------------------------------------------------------------
    * 0. CHECK THE REQUEST
    *------------------------------------------------------------------
    if "`sheet'" == "" local sheet "Tables"
    local badsheet = (length("`sheet'") > 31)
    foreach ch in 58 42 63 47 92 91 93 {
        if strpos("`sheet'", char(`ch')) local badsheet = 1
    }
    if `badsheet' {
        di as error "sheet() must be at most 31 characters and cannot contain : * ? / \ [ ]"
        exit 198
    }
    if lower("`sheet'") == "index" {
        di as error "sheet(Index) is reserved for the table index; choose another name"
        exit 198
    }
    if `decimals' < 0 | `decimals' > 4 {
        di as error "decimals() must be between 0 and 4"
        exit 198
    }
    if `strmax' < 0 {
        di as error "strmax() must be 0 or more"
        exit 198
    }

    * write .xlsx (an .xls name is switched to .xlsx)
    if regexm(lower(`"`using'"'), "\.xls$") local using `"`using'x"'
    else if !regexm(lower(`"`using'"'), "\.xlsx$") local using `"`using'.xlsx"'
    capture confirm file `"`using'"'
    if _rc == 0 & "`replace'" == "" {
        di as error `"file `using' already exists; add option {bf:replace}"'
        exit 602
    }

    local both : list categorical & continuous
    if "`both'" != "" {
        di as error "`both' cannot be in both categorical() and continuous()"
        exit 198
    }

    local pfmt "0"
    if `decimals' > 0 local pfmt = "0." + substr("0000", 1, `decimals')

    unab allvars : _all
    if "`by'" != "" local allvars : list allvars - by

    if "`varlist'" == "" local targets "`allvars'"
    else                 local targets "`varlist'"
    if "`by'" != "" local targets : list targets - by
    if "`targets'" == "" {
        di as error "no variables to export"
        exit 102
    }

    marksample touse, novarlist strok
    qui count if `touse'
    if r(N) == 0 error 2000

    *------------------------------------------------------------------
    * 1. DETECT MULTIPLE-SELECT BLOCKS (on the full data, before any
    *    sample restriction, so thin subsets do not hide a question)
    *------------------------------------------------------------------
    local nblk        = 0
    local consumed    ""
    local inddums     ""
    local indpars     ""
    local confirmedqn ""

    if "`multiselect'" == "" {

        foreach v of local allvars {

            if strpos(" `consumed' ", " `v' ") continue

            local vtype : type `v'
            local visstr = (substr("`vtype'", 1, 3) == "str")

            if `visstr' == 0 {
                qui count if !missing(`v') & `v' != int(`v')
                if r(N) > 0 continue
            }
            qui count if !missing(`v')
            if r(N) == 0 continue

            local qn ""
            local rk ""
            if regexm("`v'", "^(.+)_([0-9]+)$") {
                local qn = regexs(1)
                local rk = regexs(2)
            }

            * option sets by name: direct P_<code>, nested Q_<code>_<k>
            local dset ""
            capture unab cnd : `v'_*
            if _rc == 0 {
                foreach cd of local cnd {
                    local code = substr("`cd'", length("`v'") + 2, .)
                    if !regexm("`code'", "^_?[0-9]+$") continue
                    _et_isdummy `cd'
                    if `r(ok)' local dset "`dset' `cd'"
                }
            }

            local nset ""
            if "`qn'" != "" {
                capture unab cnn : `qn'_*_`rk'
                if _rc == 0 {
                    foreach cd of local cnn {
                        local code = substr("`cd'", length("`qn'") + 2, ///
                            length("`cd'") - length("`qn'") - length("`rk'") - 2)
                        if !regexm("`code'", "^_?[0-9]+$") continue
                        _et_isdummy `cd'
                        if `r(ok)' local nset "`nset' `cd'"
                    }
                }
            }

            local nd : word count `dset'
            local nn : word count `nset'
            if `nd' < 2 & `nn' < 2 continue

            local evid = 0
            if "`qn'" != "" {
                if strpos(" `confirmedqn' ", " `qn' ") local evid = 1
            }

            if `nn' > `nd'               local pattern "nested"
            else if `nd' > `nn'           local pattern "direct"
            else if "`qn'" != "" & `evid' local pattern "nested"
            else                          local pattern "direct"

            if "`pattern'" == "direct" local oset "`dset'"
            else                       local oset "`nset'"

            * every option must equal 1 exactly where the parent holds it
            local keep ""
            local nok  = 0
            local nbad = 0
            foreach cd of local oset {
                _et_code `cd' "`pattern'" "`v'" "`qn'" "`rk'"
                local ncode "`r(code)'"
                if `visstr' local sel `"(strpos(" " + `v' + " ", " `ncode' ") > 0)"'
                else        local sel "(`v' == `ncode')"
                qui count if (`cd' == 1) != `sel' & !missing(`v')
                if r(N) == 0 {
                    local ++nok
                    local keep "`keep' `cd'"
                }
                else local ++nbad
            }
            if `nbad' > 0 | `nok' < 2 continue

            local ordered ""
            local codes   ""
            foreach av of local allvars {
                if strpos(" `keep' ", " `av' ") {
                    _et_code `av' "`pattern'" "`v'" "`qn'" "`rk'"
                    local ordered "`ordered' `av'"
                    local codes   "`codes' `r(code)'"
                }
            }

            * A numeric parent holds one code per respondent, so it is a
            * single choice whose 0/1 dummies were generated from it
            * (tab v, gen(v_)).  The parent gets its own table and the
            * dummies none - unless a sibling repeat instance has already
            * shown this question to be a select_multiple.
            if `visstr' == 0 & `evid' == 0 {
                foreach d of local ordered {
                    local inddums "`inddums' `d'"
                    local indpars "`indpars' `v'"
                }
                local consumed = trim(itrim("`consumed' `v' `ordered'"))
                continue
            }

            local ++nblk
            local blk`nblk'_dums   = trim(itrim("`ordered'"))
            local blk`nblk'_codes  = trim(itrim("`codes'"))
            local blk`nblk'_parent "`v'"
            local blk`nblk'_stub   = cond("`pattern'" == "nested", "`qn'", "`v'")
            local blk`nblk'_pat    "`pattern'"
            local blk`nblk'_rk     "`rk'"
            local consumed = trim(itrim("`consumed' `v' `ordered'"))
            if "`pattern'" == "nested" & strpos(" `confirmedqn' ", " `qn' ") == 0 {
                local confirmedqn = trim(itrim("`confirmedqn' `qn'"))
            }
        }

        * carry a confirmed repeat question across instances that could
        * not be settled on their own
        local nb0 = `nblk'
        forvalues b = 1/`nb0' {
            if "`blk`b'_pat'" != "nested" continue
            local bqn   "`blk`b'_stub'"
            local brk   "`blk`b'_rk'"
            local bcodes ""
            foreach d of local blk`b'_dums {
                local cc = substr("`d'", length("`bqn'") + 2, ///
                    length("`d'") - length("`bqn'") - length("`brk'") - 2)
                local bcodes "`bcodes' `cc'"
            }
            capture unab sibs : `bqn'_*
            if _rc continue
            local kk ""
            foreach s of local sibs {
                if regexm("`s'", "^`bqn'_(.+)_([0-9]+)$") {
                    local k2 = regexs(2)
                    if "`k2'" != "`brk'" & strpos(" `kk' ", " `k2' ") == 0 {
                        local kk "`kk' `k2'"
                    }
                }
            }
            foreach k2 of local kk {
                capture confirm variable `bqn'_`k2', exact
                if _rc continue
                if strpos(" `consumed' ", " `bqn'_`k2' ") continue
                local dl ""
                local cl ""
                foreach cc of local bcodes {
                    capture confirm variable `bqn'_`cc'_`k2', exact
                    if _rc continue
                    if strpos(" `consumed' ", " `bqn'_`cc'_`k2' ") continue
                    _et_isdummy `bqn'_`cc'_`k2'
                    if !`r(ok)' continue
                    local dl "`dl' `bqn'_`cc'_`k2'"
                    local cl "`cl' `=subinstr("`cc'", "_", "-", 1)'"
                }
                if `: word count `dl'' < 2 continue
                local ++nblk
                local blk`nblk'_dums   = trim(itrim("`dl'"))
                local blk`nblk'_codes  = trim(itrim("`cl'"))
                local blk`nblk'_parent "`bqn'_`k2'"
                local blk`nblk'_stub   "`bqn'"
                local blk`nblk'_pat    "nested"
                local blk`nblk'_rk     "`k2'"
                local consumed = trim(itrim("`consumed' `bqn'_`k2' `dl'"))
            }
        }

        * option dummies whose parent is not in the data: group by stub
        local pstubs ""
        local np = 0
        foreach v of local allvars {
            if strpos(" `consumed' ", " `v' ") continue
            if !regexm("`v'", "^(.*[^_])_(_?[0-9]+)$") continue
            local stub = regexs(1)
            local code = subinstr(regexs(2), "_", "-", 1)
            capture confirm variable `stub', exact
            if _rc == 0 continue
            _et_isdummy `v'
            if !`r(ok)' continue
            local pos : list posof "`stub'" in pstubs
            if `pos' == 0 {
                local pstubs "`pstubs' `stub'"
                local ++np
                local pos = `np'
                local pm`pos' ""
                local pc`pos' ""
            }
            local pm`pos' "`pm`pos'' `v'"
            local pc`pos' "`pc`pos'' `code'"
        }
        forvalues p = 1/`np' {
            if `: word count `pm`p''' < 2 continue
            * q_1 q_2 q_3 can also be the instances of one question asked
            * in a repeat group (a yes/no per loan, say); those stay apart
            _et_repeats `pm`p''
            if `r(yes)' continue
            local ++nblk
            local blk`nblk'_dums   = trim(itrim("`pm`p''"))
            local blk`nblk'_codes  = trim(itrim("`pc`p''"))
            local blk`nblk'_parent ""
            local blk`nblk'_stub   : word `p' of `pstubs'
            local blk`nblk'_pat    "noparent"
            local consumed = trim(itrim("`consumed' `pm`p''"))
        }
    }
    forvalues b = 1/`nblk' {
        local blk`b'_all = trim("`blk`b'_parent' `blk`b'_dums'")
    }

    *------------------------------------------------------------------
    * 2. SAMPLE AND COLUMN GROUPS
    *------------------------------------------------------------------
    preserve
    qui keep if `touse'

    tempvar g casev
    local show = ("`by'" != "")
    local nbymiss = 0
    if `show' {
        qui egen long `g' = group(`by'), label
        qui count if missing(`g')
        local nbymiss = r(N)
        if `nbymiss' > 0 qui drop if missing(`g')
        if _N == 0 {
            di as error "every observation is missing on `by'"
            exit 2000
        }
        qui summarize `g', meanonly
        local G = r(max)
        if `G' > 1000 {
            di as error "`by' has `G' categories; at most 1,000 fit in one Excel sheet"
            exit 198
        }
        forvalues j = 1/`G' {
            local gl`j' : label (`g') `j'
        }
        local bylab : variable label `by'
        if `"`bylab'"' == "" local bylab "`by'"
        local bydesc "`by'"
        if `"`bylab'"' != "`by'" local bydesc `"`by' - `bylab'"'
    }
    else {
        qui gen byte `g' = 1
        local G = 1
    }
    local Nused = _N

    *------------------------------------------------------------------
    * 3. PLAN: ONE ENTRY PER TABLE OR SKIPPED VARIABLE
    *------------------------------------------------------------------
    local ne    = 0
    local bdone ""
    foreach v of local targets {

        local b = 0
        forvalues bb = 1/`nblk' {
            if `b' == 0 & strpos(" `blk`bb'_all' ", " `v' ") local b = `bb'
        }
        if `b' > 0 {
            if strpos(" `bdone' ", " `b' ") continue
            local bdone "`bdone' `b'"
            local ++ne
            local e`ne'_kind "multi"
            local e`ne'_var  = cond("`blk`b'_parent'" != "", "`blk`b'_parent'", "`blk`b'_stub'")
            local e`ne'_blk  `b'
            continue
        }

        local ++ne
        local e`ne'_var "`v'"

        local pos : list posof "`v'" in inddums
        if `pos' > 0 {
            local e`ne'_kind "ind"
            local e`ne'_par : word `pos' of `indpars'
            continue
        }

        local vtype : type `v'
        local vfmt  : format `v'
        local vl    : value label `v'

        if substr("`vtype'", 1, 3) == "str" {
            * Text answers.  Data typed into Excel keeps its categories as
            * text ("Male", "Lack of money" ...), so a text variable with at
            * most strmax() different answers gets a table like any single
            * choice.  Free text, names, IDs, dates written as text and form
            * metadata do not; categorical() forces a table.
            mata: _et_strinfo("`v'")
            local forced = strpos(" `categorical' ", " `v' ") > 0
            if `nnm' == 0 {
                local e`ne'_kind "empty"
                continue
            }
            local asnumber = 0
            if !`forced' {
                * (inlist() takes at most 10 text arguments)
                local lv = lower("`v'")
                if inlist("`lv'", "key", "instanceid", "instancename", ///
                          "submissiondate", "starttime", "endtime") | ///
                   inlist("`lv'", "deviceid", "subscriberid", "simid", ///
                          "devicephonenum", "username", "caseid") | ///
                   inlist("`lv'", "formdef_version", "audit", "text_audit", ///
                          "today", "phonenumber") {
                    local e`ne'_kind "meta"
                    continue
                }
                if `ndate' {
                    local e`ne'_kind "date"
                    continue
                }
                * Numbers stored as text (Excel often does this): with more
                * than 10 different values it is a measurement such as age
                * or income, converted and handled below like any number;
                * a few values (a 1 to 5 rating) are categories.
                if `nall' & `nun' > 10 {
                    tempvar num
                    local vlab : variable label `v'
                    qui gen double `num' = real(`v')
                    drop `v'
                    rename `num' `v'
                    label variable `v' `"`vlab'"'
                    local asnumber = 1
                }
                else if `nun' > `strmax' {
                    local e`ne'_kind "string"
                    local e`ne'_why "No table: text with `nun' different answers (more than `strmax')"
                    continue
                }
                else if `nun' == `nnm' & `nnm' >= 10 {
                    local e`ne'_kind "string"
                    local e`ne'_why "No table: text, every answer different"
                    continue
                }
            }
            local e`ne'_text "1"
            if !`asnumber' {
                * recode the text into a labelled number under the same
                * name (preserve brings the text back at the end) and
                * tabulate that like any single choice
                tempvar enc
                local vlab : variable label `v'
                mata: _et_strcat("`v'", "`enc'")
                drop `v'
                rename `enc' `v'
                label values `v' `enc'
                label variable `v' `"`vlab'"'
                local e`ne'_kind "cat"
                continue
            }
            local vfmt : format `v'
            local vl ""
        }
        qui count if !missing(`v')
        if r(N) == 0 {
            local e`ne'_kind "empty"
            continue
        }
        if strpos(" `continuous' ", " `v' ") {
            local e`ne'_kind "cont"
            continue
        }
        if strpos(" `categorical' ", " `v' ") {
            local e`ne'_kind "cat"
            continue
        }
        if regexm("`vfmt'", "^%-?(t|d)") {
            local e`ne'_kind "date"
            continue
        }
        * identifiers: an id-like name (UID, hhid, resp_id, key ...) and a
        * different value in every row
        if regexm(lower("`v'"), "id$|(^|_)(key|uuid|serial|sl|slno)$|^(id|key|uuid)_") {
            capture isid `v', missok
            if _rc == 0 {
                local e`ne'_kind "id"
                continue
            }
        }
        * form metadata that SurveyCTO, ODK and Kobo add to every export
        if inlist(lower("`v'"), "formdef_version", "deviceid", "subscriberid", ///
                  "simid", "devicephonenum", "caseid", "instanceid", "key") {
            local e`ne'_kind "meta"
            continue
        }
        * the parts of a GPS reading (q_latitude, qLongitude ...)
        if regexm(lower("`v'"), "(latitude|longitude|altitude|accuracy)$") {
            local e`ne'_kind "gps"
            continue
        }
        * phone numbers: a phone-like name and values of eight digits or more
        if regexm(lower("`v'"), "phone|mobile|contact|cell") {
            qui summarize `v', meanonly
            if r(min) >= 1e7 {
                local e`ne'_kind "phone"
                continue
            }
        }
        * the source of generated 0/1 indicators is a set of categories
        if strpos(" `indpars' ", " `v' ") {
            local e`ne'_kind "cat"
            continue
        }
        if "`vl'" != "" {
            * a value label that only names a few special codes (-99 Don't
            * know) on an otherwise continuous variable does not make it
            * categorical
            mata: st_local("labv", _et_vlvals("`vl'"))
            mata: st_local("nun", strofreal(_et_nunlab("`v'", "`labv'")))
            if `nun' > 10 local e`ne'_kind "cont"
            else          local e`ne'_kind "cat"
            continue
        }
        qui count if !missing(`v') & !inlist(`v', 0, 1)
        if r(N) == 0 local e`ne'_kind "cat"
        else         local e`ne'_kind "cont"
    }

    local ntab = 0
    forvalues i = 1/`ne' {
        if inlist("`e`i'_kind'", "cat", "multi", "cont") local ++ntab
    }

    *------------------------------------------------------------------
    * 4. WRITE THE TABLES
    *------------------------------------------------------------------
    di as text _n "{hline 72}"
    di as text "{bf:exporttables}  " as result "`ntab'" as text " tables" ///
        cond(`show', " by `by' (`G' " + cond(`G' == 1, "category", "categories") + " + Total)", " (one-way)")
    di as text "  file : " as result `"`using'"'
    if `nbymiss' > 0 {
        di as text "  note : " as result "`nbymiss'" as text ///
            " observations with missing `by' are left out of every table"
    }
    di as text "{hline 72}"

    * the workbook is assembled in a scratch folder and zipped at the end
    capture confirm file `"`using'"'
    if _rc == 0 {
        capture erase `"`using'"'
        if _rc {
            di as error `"cannot replace `using'; is it open in Excel? Close it and run again"'
            exit 608
        }
    }
    local K = cond(`show', `G' + 1, 1)
    local maxcol = 3
    forvalues i = 1/`ne' {
        if inlist("`e`i'_kind'", "cat", "multi") local maxcol = max(`maxcol', 1 + 2 * `K')
        if "`e`i'_kind'" == "cont"               local maxcol = max(`maxcol', 1 + `K')
    }
    tempfile scratch
    local xdir "`scratch'_xlsx"
    mkdir "`xdir'"
    mata: _et_open("`xdir'", "`sheet'", `maxcol', "`pfmt'")

    tempname T U
    local row     = 1
    local tno     = 0
    local nok     = 0
    local nfail   = 0
    local nstr    = 0
    local nskip   = 0
    local ncat    = 0
    local nmul    = 0
    local ncon    = 0
    local ntxt    = 0
    local wdone   = 0

    forvalues i = 1/`ne' {

        local kind "`e`i'_kind'"
        local v    "`e`i'_var'"
        local e`i'_row ""
        local e`i'_tno ""

        * ---------- no table for these ----------
        if !inlist("`kind'", "cat", "multi", "cont") {
            if "`kind'" == "string" {
                local ++nstr
                local e`i'_res "`e`i'_why'"
                if "`e`i'_res'" == "" local e`i'_res "No table: text variable"
            }
            else if "`kind'" == "date" {
                local ++nskip
                local e`i'_res "No table: date / time variable"
            }
            else if "`kind'" == "id" {
                local ++nskip
                local e`i'_res "No table: identifier (unique in every row)"
            }
            else if "`kind'" == "ind" {
                local ++nskip
                local e`i'_res "No table: 0/1 indicator of `e`i'_par' (see its table)"
            }
            else if "`kind'" == "meta" {
                local ++nskip
                local e`i'_res "No table: form metadata"
            }
            else if "`kind'" == "gps" {
                local ++nskip
                local e`i'_res "No table: GPS reading"
            }
            else if "`kind'" == "phone" {
                local ++nskip
                local e`i'_res "No table: phone number"
            }
            else {
                local ++nskip
                local e`i'_res "No table: all values missing"
            }
            local ktxt = cond("`kind'" == "ind", "indicator", ///
                         cond("`kind'" == "id", "identifier", ///
                         cond("`kind'" == "meta", "metadata", ///
                         cond("`kind'" == "string", "text", "`kind'"))))
            di as text "  " _col(14) %-30s abbrev("`v'", 30) %-19s "`ktxt'" ///
                as text "skipped"
            continue
        }

        local ++tno
        local ++wdone
        local start = `row'
        local endrow = `row'

        local tlab ""
        if "`kind'" != "multi" local tlab : variable label `v'
        if "`kind'" == "multi" {
            local b      = `e`i'_blk'
            local par    "`blk`b'_parent'"
            local dums   "`blk`b'_dums'"
            local codes  "`blk`b'_codes'"
            local k : word count `dums'
            local tlab ""
            if "`par'" != "" local tlab : variable label `par'
            local shown "`v' (`k' options)"
        }
        else local shown "`v'"

        di as text "  " %-11s "[`wdone'/`ntab']" ///
            as result %-30s abbrev("`shown'", 30) _continue
        local ktxt = cond("`kind'" == "cat", "single choice", ///
                     cond("`kind'" == "multi", "multiple choice", "continuous"))
        if "`e`i'_text'" == "1" {
            local ktxt = cond("`kind'" == "cat", "single (text)", "continuous (text)")
        }
        di as text %-19s "`ktxt'" _continue

        capture noisily _et_one , kind(`kind') v(`v') g(`g') ggroups(`G') ///
            show(`show') row(`row') tno(`tno') pfmt(`pfmt') dec(`decimals') ///
            tlab(`"`tlab'"') `allcats' par(`par') dums(`dums') codes(`codes') ///
            bylab(`"`bylab'"') tmat(`T') umat(`U') casev(`casev')
        local rc = _rc

        if `rc' == 0 {
            local endrow = r(endrow)
            local ++nok
            if "`kind'" == "cat"   local ++ncat
            if "`e`i'_text'" == "1" local ++ntxt
            if "`kind'" == "multi" local ++nmul
            if "`kind'" == "cont"  local ++ncon
            if "`kind'" == "multi" & `"`r(title)'"' != "" {
                local e`i'_lab `"`r(title)'"'
            }
            local e`i'_res "Exported"
            local e`i'_row `start'
            local e`i'_tno `tno'
            local row = `endrow' + 3
            di as result "done"
        }
        else {
            local ++nfail
            local e`i'_res "FAILED (error `rc')"
            local failtxt `"Table `tno': `v' could not be exported (Stata error `rc')"'
            capture mata: _et_fail(`start', "failtxt")
            local row = max(`row', `endrow') + 3
            di as error "FAILED (error `rc')"
        }
        capture drop `casev'
    }

    *------------------------------------------------------------------
    * 5. INDEX SHEET, THEN SAVE THE WORKBOOK
    *------------------------------------------------------------------
    local dname = c(filename)
    if `"`dname'"' == "" local dname "(data in memory, not saved)"
    local created = trim("`c(current_date)'") + " `c(current_time)'"
    if `show' local cols `"`bydesc' (`G' categories + Total)"'
    else      local cols "None (one-way tables)"
    forvalues i = 1/`ne' {
        if `"`e`i'_lab'"' == "" & "`e`i'_kind'" != "multi" {
            local e`i'_lab : variable label `e`i'_var'
        }
    }

    capture noisily mata: _et_index(`ne')
    if _rc di as error "  the Index sheet could not be completed (error `=_rc')"

    local rc = 0
    capture noisily mata: _et_close()
    if _rc == 0 {
        local pwd `"`c(pwd)'"'
        quietly cd "`xdir'"
        capture noisily quietly zipfile "[Content_Types].xml" _rels xl, ///
            saving("`scratch'_book.zip", replace)
        local rc = _rc
        quietly cd `"`pwd'"'
        if `rc' == 0 capture copy "`scratch'_book.zip" `"`using'"', replace
        if `rc' == 0 local rc = _rc
    }
    else local rc = _rc
    capture erase "`scratch'_book.zip"
    capture mata: _et_clean("`xdir'")
    if `rc' {
        di as error `"  the workbook could not be saved to `using' (error `rc')"'
        exit `rc'
    }

    restore

    *------------------------------------------------------------------
    * 6. REPORT
    *------------------------------------------------------------------
    * a full path, so the link opens the file wherever Stata's folder is
    local usefull `"`using'"'
    if !(substr(`"`using'"', 2, 1) == ":" | inlist(substr(`"`using'"', 1, 1), "/", "\", "~")) {
        local usefull `"`c(pwd)'`c(dirsep)'`using'"'
    }

    di as text "{hline 72}"
    di as text "  Tables exported      : " as result `nok' ///
        as text "  (single `ncat', multiple `nmul', continuous `ncon')"
    if `ntxt' > 0 {
        di as text "  From text variables  : " as result `ntxt' ///
            as text "  (categories, or numbers stored as text)"
    }
    di as text "  Text variables       : " as result `nstr' ///
        as text "  (no table: more than `strmax' different answers, or every answer different)"
    if `nskip' > 0 {
        di as text "  Other skipped        : " as result `nskip' ///
            as text "  (identifier, date, metadata, GPS, phone, 0/1 indicator or all missing)"
    }
    if `nfail' > 0 {
        di as error "  Failed               : `nfail'  (see the Index sheet)"
    }
    di as text "  Saved to             : " as result `"`using'"'
    di as text `"  Open                 : {browse `"`usefull'"':click to open}"'
    di as text "{hline 72}"

    return scalar N_tables   = `nok'
    return scalar N_single   = `ncat'
    return scalar N_multiple = `nmul'
    return scalar N_cont     = `ncon'
    return scalar N_string   = `nstr'
    return scalar N_text     = `ntxt'
    return scalar N_skipped  = `nskip'
    return scalar N_failed   = `nfail'
    return scalar N          = `Nused'
    return local  using        `"`using'"'
    return local  by           "`by'"
end


*======================================================================
* _et_one: compute and write one table, starting at row()
*======================================================================
program define _et_one, rclass
    syntax , kind(string) v(string) g(string) ggroups(integer)          ///
        show(integer) row(integer) tno(integer) pfmt(string)            ///
        dec(integer) tmat(string) umat(string) casev(string)            ///
        [ tlab(string) allcats par(string) dums(string)                 ///
          codes(string) bylab(string) ]

    local G = `ggroups'
    local r = `row'

    if `"`tlab'"' == "" local tlab "`v'"

    *-------------------------------------------------- compute
    if "`kind'" == "cat" {
        local vl : value label `v'
        qui count if !missing(`v') & !inlist(`v', 0, 1)
        local yesno = (r(N) == 0)
        local extra ""
        if "`allcats'" != "" & "`vl'" != "" {
            mata: st_local("extra", _et_vlvals("`vl'"))
        }
        mata: _et_cat("`v'", "`g'", `G', `show', "`extra'", `dec', "`tmat'", "`umat'")
        local L = rowsof(`umat')
        forvalues l = 1/`L' {
            local val = `umat'[`l', 1]
            local rl`l' = string(`val', "%18.0g")
            if "`vl'" != "" & `val' == int(`val') {
                local lb : label `vl' `val', strict
                if `"`lb'"' != "" local rl`l' `"`lb'"'
            }
            else if "`vl'" == "" & `yesno' {
                local rl`l' = cond(`val' == 1, "Yes", "No")
            }
        }
        local ++L
        local rl`L' "Total"
        local hdr "Response"
    }
    else if "`kind'" == "multi" {
        * option labels: the dummy's variable label, else the parent's
        * value label, else the code; a shared "Question: " prefix on
        * every dummy label is lifted into the title
        local k : word count `dums'
        local pvl ""
        if "`par'" != "" {
            capture confirm numeric variable `par'
            if _rc == 0 local pvl : value label `par'
        }
        local pref ""
        local same = 1
        forvalues l = 1/`k' {
            local d : word `l' of `dums'
            local dl : variable label `d'
            local p0 ""
            if strpos(`"`dl'"', ":") local p0 = strtrim(substr(`"`dl'"', 1, strpos(`"`dl'"', ":") - 1))
            if `l' == 1 local pref `"`p0'"'
            else if `"`p0'"' != `"`pref'"' local same = 0
        }
        if `"`pref'"' == "" local same = 0

        forvalues l = 1/`k' {
            local d  : word `l' of `dums'
            local cd : word `l' of `codes'
            local dl : variable label `d'
            if `same' local dl = strtrim(substr(`"`dl'"', strpos(`"`dl'"', ":") + 1, .))
            if `"`tlab'"' != "" & `"`tlab'"' != "`v'" & strpos(`"`dl'"', `"`tlab'"') == 1 {
                local dl = strtrim(substr(`"`dl'"', length(`"`tlab'"') + 1, .))
                local dl = strtrim(regexr(`"`dl'"', "^[:/ -]+", ""))
            }
            if `"`dl'"' == "" & "`pvl'" != "" {
                capture local dl : label `pvl' `cd', strict
            }
            if `"`dl'"' == "" local dl "Option `cd'"
            local rl`l' `"`dl'"'
        }
        if (`"`tlab'"' == "" | `"`tlab'"' == "`v'") & `same' local tlab `"`pref'"'

        * valid cases: answered the question (parent non-missing), or,
        * with no parent, non-missing on at least one option
        if "`par'" != "" {
            qui gen byte `casev' = !missing(`par')
        }
        else {
            qui egen byte `casev' = rownonmiss(`dums')
            qui replace `casev' = `casev' > 0
        }
        mata: _et_multi("`dums'", "`casev'", "`g'", `G', `show', `dec', "`tmat'")
        local L = `k' + 1
        local rl`L' "Valid cases (N)"
        local hdr "Response"
    }
    else {
        * labelled special codes (-99 Don't know ...) are left out
        local vl : value label `v'
        local excl ""
        if "`vl'" != "" mata: st_local("excl", _et_vlvals("`vl'"))
        mata: _et_cont("`v'", "`g'", `G', `show', "`excl'", "`tmat'")
        local L = 7
        local i = 0
        foreach s in "N" "Mean" "Median" "Mode" "SD" "Min" "Max" {
            local ++i
            local rl`i' "`s'"
        }
        local hdr "Statistic"
        local exnote ""
        foreach x of local excl {
            qui count if `v' == `x'
            if r(N) > 0 {
                local nx = r(N)
                local lb : label `vl' `x', strict
                if `"`exnote'"' != "" local exnote `"`exnote'; "'
                local exnote `"`exnote'`x' `lb' (n=`nx')"'
            }
        }
    }

    *-------------------------------------------------- write
    * the Mata writer reads ttl, hdr, note and rl1..rlL from here and
    * sets endrow and lastcol
    local star = cond("`kind'" == "multi", " *", "")
    local ttl `"Table `tno': `tlab' [`v']`star'"'
    if `show' local ttl `"`ttl' - by `bylab'"'

    local note ""
    if "`kind'" == "multi" {
        local note "* Multiple-response question: N = valid cases (respondents who answered the question). Percentages are of valid cases and can add up to more than 100%."
    }
    if "`kind'" == "cont" & `"`exnote'"' != "" {
        local note `"Not included in the statistics: `exnote'"'
    }

    mata: _et_write("`kind'", `show', `G', "`g'", `row', `L', "`tmat'", "`pfmt'")

    return scalar endrow  = `endrow'
    return scalar lastcol = `lastcol'
    return local  title     `"`tlab'"'
end


*======================================================================
* small helpers
*======================================================================

* is a numeric 0/1 variable with at least one non-missing value?
program define _et_isdummy, rclass
    args v
    return scalar ok = 0
    capture confirm numeric variable `v'
    if _rc exit
    qui count if !missing(`v')
    if r(N) == 0 exit
    qui count if !inlist(`v', 0, 1) & !missing(`v')
    if r(N) > 0 exit
    return scalar ok = 1
end

* do these 0/1 variables look like the instances of one question asked in a
* repeat group, rather than the options of one multiple-choice question?
* Instances share their wording ("Type of policy #4", "#5" ...) or carry a
* value label naming more than 0 and 1; options are worded differently.
program define _et_repeats, rclass
    syntax varlist
    return scalar yes = 0

    local first : word 1 of `varlist'
    local vl1 : value label `first'
    local samevl = ("`vl1'" != "")
    if `samevl' {
        mata: st_local("labv", _et_vlvals("`vl1'"))
        local binary = 1
        foreach x of local labv {
            if !inlist(`x', 0, 1) local binary = 0
        }
        if `binary' local samevl = 0
    }

    local samelb = 1
    local l1 ""
    local i = 0
    foreach v of local varlist {
        local ++i
        local vl : value label `v'
        if "`vl'" != "`vl1'" local samevl = 0
        local lb : variable label `v'
        local lb = ustrregexra(lower(`"`lb'"'), "\s*#\s*[0-9]+", "")
        local lb = ustrregexra(`"`lb'"', "\s+[0-9]+\s*$", "")
        local lb = strtrim(stritrim(`"`lb'"'))
        if `i' == 1 local l1 `"`lb'"'
        if `"`lb'"' == "" | `"`lb'"' != `"`l1'"' local samelb = 0
    }
    if `samevl' | `samelb' return scalar yes = 1
end

* option code carried by a dummy's name (q5__99 -> -99)
program define _et_code, rclass
    args cd pattern v qn rk
    if "`pattern'" == "direct" {
        local code = substr("`cd'", length("`v'") + 2, .)
    }
    else {
        local code = substr("`cd'", length("`qn'") + 2, ///
            length("`cd'") - length("`qn'") - length("`rk'") - 2)
    }
    return local code = subinstr("`code'", "_", "-", 1)
end


*======================================================================
* Mata: counting and statistics
*======================================================================
version 16
mata:

real scalar _et_find(real colvector u, real scalar x)
{
    real scalar lo, hi, mid
    lo = 1
    hi = rows(u)
    while (lo <= hi) {
        mid = floor((lo + hi) / 2)
        if (u[mid] == x) return(mid)
        if (u[mid] < x) lo = mid + 1
        else hi = mid - 1
    }
    return(0)
}

/* counts C (rows x G) and bases (1 x G) -> [N % N % ... N_total %_total],
   with the base as an extra last row                                    */
void _et_pack(real matrix C, real rowvector base, real scalar show,
              real scalar dec, real scalar basepct, string scalar mT)
{
    real matrix A, T
    real rowvector b
    real scalar j, L, K

    A = C, rowsum(C)
    b = base, rowsum(base)
    L = rows(A)
    K = cols(A)
    T = J(L + 1, 2 * K, .)
    for (j = 1; j <= K; j++) {
        if (L > 0) T[|1, 2*j-1 \ L, 2*j-1|] = A[., j]
        T[L + 1, 2*j-1] = b[j]
        if (b[j] > 0) {
            if (L > 0) T[|1, 2*j \ L, 2*j|] = round(100 :* A[., j] :/ b[j], 10^(-dec))
            if (basepct) T[L + 1, 2*j] = 100
        }
    }
    if (!show) T = T[., (2*K-1, 2*K)]
    st_matrix(mT, T)
}

void _et_cat(string scalar vn, string scalar gn, real scalar G,
             real scalar show, string scalar extra, real scalar dec,
             string scalar mT, string scalar mU)
{
    real colvector v, g, u
    real matrix C
    real scalar i, l

    v = st_data(., vn)
    g = st_data(., gn)
    u = select(v, v :< .)
    if (strtrim(extra) != "") u = u \ strtoreal(tokens(extra))'
    u = uniqrows(u)
    C = J(rows(u), G, 0)
    for (i = 1; i <= rows(v); i++) {
        if (missing(v[i]) | missing(g[i])) continue
        l = _et_find(u, v[i])
        if (l) C[l, g[i]] = C[l, g[i]] + 1
    }
    _et_pack(C, colsum(C), show, dec, 1, mT)
    st_matrix(mU, u)
}

void _et_multi(string scalar dums, string scalar cn, string scalar gn,
               real scalar G, real scalar show, real scalar dec,
               string scalar mT)
{
    real matrix X, C
    real colvector c, g
    real rowvector base
    real scalar i

    X = st_data(., tokens(dums))
    c = st_data(., cn)
    g = st_data(., gn)
    C = J(cols(X), G, 0)
    base = J(1, G, 0)
    for (i = 1; i <= rows(X); i++) {
        if (c[i] != 1 | missing(g[i])) continue
        base[g[i]] = base[g[i]] + 1
        C[., g[i]] = C[., g[i]] + (X[i, .] :== 1)'
    }
    _et_pack(C, base, show, dec, 0, mT)
}

real colvector _et_stats(real colvector x)
{
    real colvector s, r
    real scalar n, i, run, best, bestn

    r = J(7, 1, .)
    n = rows(x)
    r[1] = n
    if (n == 0) return(r)
    s = sort(x, 1)
    r[2] = mean(s)
    if (mod(n, 2)) r[3] = s[(n + 1) / 2]
    else           r[3] = (s[n / 2] + s[n / 2 + 1]) / 2
    best  = s[1]
    bestn = 1
    run   = 1
    for (i = 2; i <= n; i++) {
        if (s[i] == s[i - 1]) run = run + 1
        else run = 1
        if (run > bestn) {
            bestn = run
            best  = s[i]
        }
    }
    /* with no value occurring twice there is no mode: leave it blank
       rather than report the smallest value */
    if (bestn > 1 | n == 1) r[4] = best
    if (n > 1) r[5] = sqrt(quadvariance(s))
    r[6] = s[1]
    r[7] = s[n]
    return(r)
}

void _et_cont(string scalar vn, string scalar gn, real scalar G,
              real scalar show, string scalar excl, string scalar mT)
{
    real colvector v, g, keep, ex
    real matrix T
    real scalar j, K

    v = st_data(., vn)
    g = st_data(., gn)
    keep = (v :< .)
    if (strtrim(excl) != "") {
        ex = strtoreal(tokens(excl))'
        for (j = 1; j <= rows(ex); j++) keep = keep :& (v :!= ex[j])
    }
    K = (show ? G + 1 : 1)
    T = J(7, K, .)
    for (j = 1; j <= K; j++) {
        if (show & j <= G) T[., j] = _et_stats(select(v, keep :& (g :== j)))
        else               T[., j] = _et_stats(select(v, keep))
    }
    st_matrix(mT, T)
}

/* values defined in a value label, space separated */
string scalar _et_vlvals(string scalar lbl)
{
    real colvector vals
    string colvector txt

    if (lbl == "") return("")
    if (!st_vlexists(lbl)) return("")
    st_vlload(lbl, vals, txt)
    if (rows(vals) == 0) return("")
    return(invtokens(strofreal(vals', "%18.0g")))
}

/* distinct non-missing values of a variable that carry no label */
real scalar _et_nunlab(string scalar vn, string scalar labv)
{
    real colvector u, lv
    real scalar i, n

    u = st_data(., vn)
    u = uniqrows(select(u, u :< .))
    if (strtrim(labv) == "") return(rows(u))
    lv = strtoreal(tokens(labv))'
    n = 0
    for (i = 1; i <= rows(u); i++) if (!anyof(lv, u[i])) n++
    return(n)
}

/* does a text value look like a date or a date-time?  (2026-09-01,
   01/09/2026, Sep 1, 2026 8:00:00 AM, 01Sep2026 ...)                 */
real scalar _et_datelike(string scalar x)
{
    string scalar m, y

    m = "(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*"
    y = strlower(strtrim(x))
    if (regexm(y, "^[0-9][0-9][0-9][0-9]-[0-9][0-9]?-[0-9][0-9]?")) return(1)
    if (regexm(y, "^[0-9][0-9]?[/.-][0-9][0-9]?[/.-][0-9][0-9][0-9][0-9]")) return(1)
    if (regexm(y, "^" + m + " [0-9][0-9]?,? [0-9][0-9][0-9][0-9]")) return(1)
    if (regexm(y, "^[0-9][0-9]?[ -]?" + m + "[ -]?[0-9][0-9][0-9][0-9]")) return(1)
    return(0)
}

/* for a text variable, set in the caller: nnm (non-empty values), nun
   (different values) and ndate (1 if most values are dates)          */
void _et_strinfo(string scalar vn)
{
    string colvector s, u
    real scalar      i, nd

    s = st_sdata(., vn)
    s = select(s, s :!= "")
    st_local("nnm", strofreal(rows(s)))
    st_local("nun", "0")
    st_local("nall", "0")
    st_local("ndate", "0")
    if (rows(s) == 0) return

    u = uniqrows(s)
    st_local("nun", strofreal(rows(u)))
    st_local("nall", (hasmissing(strtoreal(u)) ? "0" : "1"))
    nd = 0
    for (i = 1; i <= rows(u); i++) nd = nd + _et_datelike(u[i])
    if (nd >= 0.8 * rows(u)) st_local("ndate", "1")
}

/* sort key for natural order: case ignored and every run of digits
   padded, so "Type 2" comes before "Type 10"                         */
string scalar _et_natkey(string scalar s)
{
    string scalar out, run, c
    real scalar   i

    out = ""
    run = ""
    for (i = 1; i <= strlen(s); i++) {
        c = substr(s, i, 1)
        if (c >= "0" & c <= "9") {
            run = run + c
            continue
        }
        if (run != "") out = out + substr("000000000000", 1, max((0, 12 - strlen(run)))) + run
        run = ""
        out = out + c
    }
    if (run != "") out = out + substr("000000000000", 1, max((0, 12 - strlen(run)))) + run
    return(strlower(out))
}

/* recode text variable vn into a new numeric variable nv, 1 2 3 ... in
   natural order (numeric order when every answer is a number), with a
   value label of the same name holding the text                      */
void _et_strcat(string scalar vn, string scalar nv)
{
    string colvector s, u, key
    real colvector   x, v
    real scalar      i, k
    transmorphic     A

    s = st_sdata(., vn)
    u = uniqrows(select(s, s :!= ""))
    x = strtoreal(u)
    if (rows(u) > 1 & !hasmissing(x)) u = u[order(x, 1)]
    else if (rows(u) > 1) {
        key = J(rows(u), 1, "")
        for (k = 1; k <= rows(u); k++) key[k] = _et_natkey(u[k])
        u = u[order((key, u), (1, 2))]
    }

    A = asarray_create()
    for (k = 1; k <= rows(u); k++) asarray(A, u[k], k)
    v = J(rows(s), 1, .)
    for (i = 1; i <= rows(s); i++) {
        if (s[i] != "") v[i] = asarray(A, s[i])
    }
    (void) st_addvar("long", nv)
    st_store(., nv, v)

    /* a value label holds at most 32,000 characters per value */
    for (k = 1; k <= rows(u); k++) {
        if (strlen(u[k]) > 32000) u[k] = usubstr(u[k], 1, 30000)
    }
    if (rows(u)) st_vlmodify(nv, (1::rows(u)), u)
}

/*----------------------------------------------------------------------
  Writing the workbook.

  The .xlsx is written directly as XML and zipped with -zipfile-.  Stata's
  own Excel writers (putexcel, xl()) create a new cell style for every
  cell they format, so a formatted 64-district export slows to a crawl
  and can pass Excel's limit of 64,000 styles.  Here every cell points at
  one of a fixed set of shared styles:

     0 default          4 label, bordered       9 total %
     1 table title      5 N        #,##0       10 statistic  #,##0.00
     2 header, centred  6 %        0.0         11 footnote, italic
     3 header, left     7 total label (bold)   12 link
                        8 total N              13 sheet title
                                               14 bold
                                               15 number, left aligned

  Sheet 1 is the Index, sheet 2 the tables.  The tables sheet is streamed
  row by row as each table is finished; everything else is written by
  _et_close().
----------------------------------------------------------------------*/

void _et_set(string scalar name, transmorphic x)
{
    pointer() scalar p

    if ((p = findexternal(name)) == NULL) p = crexternal(name)
    *p = x
}

transmorphic _et_get(string scalar name)
{
    return(*findexternal(name))
}

string scalar _et_esc(string scalar s)
{
    real scalar i

    s = subinstr(s, "&", "&amp;")
    s = subinstr(s, "<", "&lt;")
    s = subinstr(s, ">", "&gt;")
    s = subinstr(s, `"""', "&quot;")
    for (i = 1; i <= 31; i++) {
        if (i != 9 & i != 10 & i != 13) s = subinstr(s, char(i), "")
    }
    return(s)
}

/* characters a number takes in #,##0 (d = 0) or #,##0.00 (d = 2) */
real scalar _et_numw(real scalar x, real scalar d)
{
    if (missing(x)) return(0)
    return(strlen(strtrim(strofreal(x, "%32." + strofreal(d) + "fc"))))
}

string scalar _et_ref(real scalar r, real scalar c)
{
    return(numtobase26(c) + strofreal(r))
}

/* string, number and empty cells */
string scalar _et_cs(real scalar r, real scalar c, real scalar st, string scalar s)
{
    if (s == "") return(_et_ce(r, c, st))
    return(`"<c r=""' + _et_ref(r, c) + `"" s=""' + strofreal(st) +
        `"" t="inlineStr"><is><t xml:space="preserve">"' + _et_esc(s) +
        "</t></is></c>")
}

string scalar _et_cn(real scalar r, real scalar c, real scalar st, real scalar x)
{
    if (missing(x)) return(_et_ce(r, c, st))
    return(`"<c r=""' + _et_ref(r, c) + `"" s=""' + strofreal(st) + `""><v>"' +
        strtrim(strofreal(x, "%21.15g")) + "</v></c>")
}

string scalar _et_ce(real scalar r, real scalar c, real scalar st)
{
    return(`"<c r=""' + _et_ref(r, c) + `"" s=""' + strofreal(st) + `""/>"')
}

string scalar _et_row(real scalar r, string scalar cells)
{
    return(`"<row r=""' + strofreal(r) + `"">"' + cells + "</row>")
}

void _et_put(string scalar file, string colvector lines)
{
    real scalar fh, i

    fh = fopen(file, "w")
    for (i = 1; i <= rows(lines); i++) fput(fh, lines[i])
    fclose(fh)
}

void _et_open(string scalar dir, string scalar sheet, real scalar maxcol,
              string scalar pfmt)
{
    real scalar fh

    mkdir(dir + "/_rels")
    mkdir(dir + "/xl")
    mkdir(dir + "/xl/_rels")
    mkdir(dir + "/xl/worksheets")

    _et_set("_et_D", dir)
    _et_set("_et_S", sheet)
    _et_set("_et_P", pfmt)
    _et_set("_et_M", J(0, 1, ""))

    /* the rows are streamed to a scratch file; _et_close() wraps them in
       the sheet once every table is in, because only then are the column
       widths known (a fixed width shows a large number as #####)        */
    _et_set("_et_W", J(1, max((maxcol, 1)), 0))
    fh = fopen(dir + "/xl/worksheets/rows.tmp", "w")
    _et_set("_et_F", fh)
}

void _et_write(string scalar kind, real scalar show, real scalar G,
               string scalar gn, real scalar r, real scalar L,
               string scalar mT, string scalar pfmt)
{
    real matrix      T
    real rowvector   W
    string rowvector glab
    string colvector M
    string scalar    vl, pct, cells, note, lab
    real scalar      fh, cont, K, lastcol, hrows, h1, h2, d1, j, l, rr,
                     tot, fr, sN, sP, sL

    fh = _et_get("_et_F")
    M  = _et_get("_et_M")
    W  = _et_get("_et_W")
    T  = st_matrix(mT)

    cont    = (kind == "cont")
    K       = (show ? G + 1 : 1)
    lastcol = (cont ? 1 + K : 1 + 2 * K)
    hrows   = ((!cont & show) ? 2 : 1)
    h1 = r + 1
    h2 = r + hrows
    d1 = h2 + 1

    glab = J(1, K, "")
    if (show) {
        vl = st_varvaluelabel(gn)
        for (j = 1; j <= G; j++) {
            if (vl != "") glab[j] = st_vlmap(vl, j)
            if (glab[j] == "") glab[j] = strofreal(j)
        }
        glab[K] = "Total"
    }

    /* title */
    fput(fh, _et_row(r, _et_cs(r, 1, 1, st_local("ttl"))))

    /* header */
    pct = (kind == "multi" ? "% of cases" : "%")
    cells = _et_cs(h1, 1, 3, st_local("hdr"))
    if (cont) {
        for (j = 1; j <= K; j++) {
            cells = cells + _et_cs(h1, j + 1, 2, (show ? glab[j] : "Value"))
            if (show) W[j + 1] = max((W[j + 1], ustrlen(glab[j]) + 1))
        }
        fput(fh, _et_row(h1, cells))
    }
    else if (hrows == 1) {
        for (j = 1; j <= K; j++) {
            cells = cells + _et_cs(h1, 2*j, 2, "N") + _et_cs(h1, 2*j + 1, 2, pct)
        }
        fput(fh, _et_row(h1, cells))
    }
    else {
        for (j = 1; j <= K; j++) {
            cells = cells + _et_cs(h1, 2*j, 2, glab[j]) + _et_ce(h1, 2*j + 1, 2)
            M = M \ (_et_ref(h1, 2*j) + ":" + _et_ref(h1, 2*j + 1))
        }
        fput(fh, _et_row(h1, cells))
        cells = _et_ce(h2, 1, 3)
        for (j = 1; j <= K; j++) {
            cells = cells + _et_cs(h2, 2*j, 2, "N") + _et_cs(h2, 2*j + 1, 2, pct)
        }
        fput(fh, _et_row(h2, cells))
        M = M \ (_et_ref(h1, 1) + ":" + _et_ref(h2, 1))
    }

    /* body: the last row of a frequency table is its Total / valid cases */
    for (l = 1; l <= L; l++) {
        rr  = d1 + l - 1
        tot = (!cont & l == L)
        sL  = (tot ? 7 : 4)
        lab = st_local("rl" + strofreal(l))
        cells = _et_cs(rr, 1, sL, lab)
        W[1]  = max((W[1], ustrlen(lab)))
        if (cont) {
            sN = (l == 1 ? 5 : 10)
            for (j = 1; j <= K; j++) {
                cells = cells + _et_cn(rr, j + 1, sN, T[l, j])
                W[j + 1] = max((W[j + 1], _et_numw(T[l, j], (sN == 5 ? 0 : 2))))
            }
        }
        else {
            sN = (tot ? 8 : 5)
            sP = (tot ? 9 : 6)
            for (j = 1; j <= K; j++) {
                cells = cells + _et_cn(rr, 2*j, sN, T[l, 2*j - 1]) +
                                _et_cn(rr, 2*j + 1, sP, T[l, 2*j])
                W[2*j] = max((W[2*j], _et_numw(T[l, 2*j - 1], 0)))
            }
        }
        fput(fh, _et_row(rr, cells))
    }

    /* footnote */
    fr   = d1 + L - 1
    note = st_local("note")
    if (note != "") {
        fr = fr + 1
        fput(fh, _et_row(fr, _et_cs(fr, 1, 11, note)))
    }

    _et_set("_et_M", M)
    _et_set("_et_W", W)
    st_local("endrow",  strofreal(fr))
    st_local("lastcol", strofreal(lastcol))
}

void _et_fail(real scalar r, string scalar loc)
{
    fput(_et_get("_et_F"), _et_row(r, _et_cs(r, 1, 1, st_local(loc))))
}

string scalar _et_kindtext(string scalar k)
{
    if (k == "cat")    return("Single choice")
    if (k == "multi")  return("Multiple choice")
    if (k == "cont")   return("Continuous")
    if (k == "string") return("Text")
    if (k == "date")   return("Date / time")
    if (k == "id")     return("Identifier")
    if (k == "ind")    return("0/1 indicator")
    if (k == "meta")   return("Metadata")
    if (k == "gps")    return("GPS")
    if (k == "phone")  return("Phone number")
    return("Empty")
}

/* the Index sheet: reads the e<i>_* locals of the calling program */
void _et_index(real scalar ne)
{
    string colvector X, info, links
    string scalar    dir, sh, e, tr, cells
    real scalar      i, r, hr
    real colvector   nums

    dir = _et_get("_et_D")
    sh  = _et_get("_et_S")

    X = `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"' \
        `"<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/></sheetViews>"' \
        `"<sheetFormatPr defaultRowHeight="15"/>"' \
        (`"<cols><col min="1" max="1" width="45" customWidth="1"/><col min="2" max="2" width="26" customWidth="1"/>"' +
         `"<col min="3" max="3" width="60" customWidth="1"/><col min="4" max="4" width="16" customWidth="1"/>"' +
         `"<col min="5" max="5" width="66" customWidth="1"/></cols>"') \
        "<sheetData>"

    X = X \ _et_row(1, _et_cs(1, 1, 13, "Table index"))

    info = ("Dataset" \ "Created" \ "Observations used" \ "Columns" \
            "Tables exported" \ "   single choice" \ "   multiple choice" \
            "   continuous" \ "Text variables without a table" \
            "Other variables skipped (reason in the list)" \
            "Tables that failed")
    nums = strtoreal((st_local("Nused") \ st_local("nok") \ st_local("ncat") \
            st_local("nmul") \ st_local("ncon") \ st_local("nstr") \
            st_local("nskip") \ st_local("nfail")))
    X = X \ _et_row(3, _et_cs(3, 1, 14, info[1]) + _et_cs(3, 2, 0, st_local("dname")))
    X = X \ _et_row(4, _et_cs(4, 1, 14, info[2]) + _et_cs(4, 2, 0, st_local("created")))
    X = X \ _et_row(5, _et_cs(5, 1, 14, info[3]) + _et_cn(5, 2, 15, nums[1]))
    X = X \ _et_row(6, _et_cs(6, 1, 14, info[4]) + _et_cs(6, 2, 0, st_local("cols")))
    for (i = 5; i <= 11; i++) {
        X = X \ _et_row(i + 2, _et_cs(i + 2, 1, 14, info[i]) + _et_cn(i + 2, 2, 15, nums[i - 3]))
    }

    hr = 15
    X = X \ _et_row(hr, _et_cs(hr, 1, 3, "Table") + _et_cs(hr, 2, 3, "Variable") +
        _et_cs(hr, 3, 3, "Label") + _et_cs(hr, 4, 3, "Type") + _et_cs(hr, 5, 3, "Result"))

    links = J(0, 1, "")
    for (i = 1; i <= ne; i++) {
        r  = hr + i
        e  = "e" + strofreal(i) + "_"
        tr = st_local(e + "row")
        if (tr != "") {
            cells = _et_cs(r, 1, 12, "Table " + st_local(e + "tno"))
            links = links \ (`"<hyperlink ref=""' + _et_ref(r, 1) + `"" location=""' +
                _et_esc("'" + subinstr(sh, "'", "''") + "'!A" + tr) + `"" display=""' +
                _et_esc("Table " + st_local(e + "tno")) + `""/>"')
        }
        else cells = _et_ce(r, 1, 4)
        cells = cells + _et_cs(r, 2, 4, st_local(e + "var")) +
                        _et_cs(r, 3, 4, st_local(e + "lab")) +
                        _et_cs(r, 4, 4, _et_kindtext(st_local(e + "kind")) +
                                        (st_local(e + "text") == "1" ? " (text)" : "")) +
                        _et_cs(r, 5, 4, st_local(e + "res"))
        X = X \ _et_row(r, cells)
    }
    X = X \ "</sheetData>"
    if (rows(links)) {
        X = X \ "<hyperlinks>" \ links \ "</hyperlinks>"
    }
    X = X \ "</worksheet>"
    _et_put(dir + "/xl/worksheets/sheet1.xml", X)
}

/* finish the tables sheet and write the remaining package parts */
void _et_close()
{
    string colvector M, R
    string scalar    dir, sh, pfmt, ns, cw
    real rowvector   W
    real scalar      fh, i, w

    dir  = _et_get("_et_D")
    sh   = _et_get("_et_S")
    pfmt = _et_get("_et_P")
    M    = _et_get("_et_M")
    W    = _et_get("_et_W")
    fclose(_et_get("_et_F"))

    /* column widths from the widest entry: labels 30 to 60 characters,
       numbers 11 to 24 */
    cw = ""
    for (i = 1; i <= cols(W); i++) {
        w  = (i == 1 ? min((60, max((30, W[i] + 3)))) : min((24, max((11, W[i] + 3)))))
        cw = cw + `"<col min=""' + strofreal(i) + `"" max=""' + strofreal(i) +
             `"" width=""' + strofreal(w) + `"" customWidth="1"/>"'
    }

    R  = cat(dir + "/xl/worksheets/rows.tmp")
    (void) _unlink(dir + "/xl/worksheets/rows.tmp")
    fh = fopen(dir + "/xl/worksheets/sheet2.xml", "w")
    fput(fh, `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"')
    fput(fh, `"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"')
    fput(fh, `"<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>"')
    fput(fh, `"<sheetFormatPr defaultRowHeight="15"/>"')
    fput(fh, "<cols>" + cw + "</cols>")
    fput(fh, "<sheetData>")
    for (i = 1; i <= rows(R); i++) fput(fh, R[i])
    fput(fh, "</sheetData>")
    if (rows(M)) {
        fput(fh, `"<mergeCells count=""' + strofreal(rows(M)) + `"">"')
        for (i = 1; i <= rows(M); i++) fput(fh, `"<mergeCell ref=""' + M[i] + `""/>"')
        fput(fh, "</mergeCells>")
    }
    fput(fh, "</worksheet>")
    fclose(fh)

    ns = "http://schemas.openxmlformats.org/"

    _et_put(dir + "/[Content_Types].xml", (
        `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<Types xmlns=""' + ns + `"package/2006/content-types">"' \
        `"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"' \
        `"<Default Extension="xml" ContentType="application/xml"/>"' \
        `"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"' \
        `"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"' \
        `"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"' \
        `"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"' \
        "</Types>"))

    _et_put(dir + "/_rels/.rels", (
        `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<Relationships xmlns=""' + ns + `"package/2006/relationships">"' \
        `"<Relationship Id="rId1" Type=""' + ns + `"officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"' \
        "</Relationships>"))

    _et_put(dir + "/xl/workbook.xml", (
        `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<workbook xmlns=""' + ns + `"spreadsheetml/2006/main" xmlns:r=""' + ns + `"officeDocument/2006/relationships">"' \
        `"<bookViews><workbookView activeTab="0"/></bookViews>"' \
        `"<sheets><sheet name="Index" sheetId="1" r:id="rId1"/>"' \
        (`"<sheet name=""' + _et_esc(sh) + `"" sheetId="2" r:id="rId2"/></sheets>"') \
        "</workbook>"))

    _et_put(dir + "/xl/_rels/workbook.xml.rels", (
        `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<Relationships xmlns=""' + ns + `"package/2006/relationships">"' \
        `"<Relationship Id="rId1" Type=""' + ns + `"officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"' \
        `"<Relationship Id="rId2" Type=""' + ns + `"officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>"' \
        `"<Relationship Id="rId3" Type=""' + ns + `"officeDocument/2006/relationships/styles" Target="styles.xml"/>"' \
        "</Relationships>"))

    _et_put(dir + "/xl/styles.xml", (
        `"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"' \
        `"<styleSheet xmlns=""' + ns + `"spreadsheetml/2006/main">"' \
        (`"<numFmts count="1"><numFmt numFmtId="164" formatCode=""' + pfmt + `""/></numFmts>"') \
        `"<fonts count="7">"' \
        `"<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><b/><sz val="14"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>"' \
        `"<font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>"' \
        "</fonts>" \
        `"<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>"' \
        `"<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>"' \
        `"<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>"' \
        `"<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>"' \
        `"<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>"' \
        `"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"' \
        `"<cellXfs count="16">"' \
        `"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"' \
        `"<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>"' \
        `"<xf numFmtId="0" fontId="5" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>"' \
        `"<xf numFmtId="0" fontId="5" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>"' \
        `"<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>"' \
        `"<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>"' \
        `"<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>"' \
        `"<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>"' \
        `"<xf numFmtId="3" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>"' \
        `"<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>"' \
        `"<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>"' \
        `"<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>"' \
        `"<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>"' \
        `"<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>"' \
        `"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"' \
        `"<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>"' \
        "</cellXfs>" \
        `"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"' \
        "</styleSheet>"))
}

/* remove the scratch folder */
void _et_clean(string scalar dir)
{
    string colvector f
    string scalar    d
    real scalar      i, k

    d = dir + "/xl/worksheets"
    f = dir(d, "files", "*")
    for (i = 1; i <= rows(f); i++) (void) _unlink(d + "/" + f[i])
    for (k = 1; k <= 3; k++) {
        d = (k == 1 ? dir + "/xl/_rels" : (k == 2 ? dir + "/xl" : dir + "/_rels"))
        f = dir(d, "files", "*")
        for (i = 1; i <= rows(f); i++) (void) _unlink(d + "/" + f[i])
    }
    f = dir(dir, "files", "*")
    for (i = 1; i <= rows(f); i++) (void) _unlink(dir + "/" + f[i])
    (void) _rmdir(dir + "/xl/worksheets")
    (void) _rmdir(dir + "/xl/_rels")
    (void) _rmdir(dir + "/xl")
    (void) _rmdir(dir + "/_rels")
    (void) _rmdir(dir)
}

end
