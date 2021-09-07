*! version 0.1  07sep2021  Diana Goldemberg,  diana_goldemberg@g.harvard.edu

/*------------------------------------------------------------------------------
  This ado saves as dta/csv/xls/xlsx similar results as "codebook, compact",
  useful because codebook does not store any results, only displays them.
  It does not allow yet the full syntax of codebook, notably [varlist] [in] [if].
------------------------------------------------------------------------------*/

cap program drop savecodebook
program define   savecodebook

  version 13
  
  syntax using/, [replace]
  
  quietly {
    
    
    *-------------------------------------------------------------
    * Check if `using' was correctly specified
    
    * Split using argument into r(folder) r(filename) and r(extension)
    mata: split_using(`"`using'"')
    local r_folder    = "`r(folder)'"
    local r_filename  = "`r(filename)'"
    local r_extension = "`r(extension)'"
  
    * Test that the folder for the codebook exists
    mata : st_numscalar("r(dirExist)", direxists("`r_folder'"))
    if `r(dirExist)' == 0  {
      noi di as error `"{phang}The folder [`r_folder'/] does not exist.{p_end}"'
      error 601
    }
    
    * Test if a supported extension was specified
    local extension_allowed = inlist("`r_extension'", ".dta", ".csv", ".xlsx", ".xls")
    if `extension_allowed' != 1 {
    	noi di as error `"The codebook may only have the file extension [.dta], [.csv], [.xslx] or [.xls]. The format [`r_extension'] is not allowed."'
      error 601
    }    
    *-------------------------------------------------------------
    
    * Preserve the original data to restore at the end
    preserve
    
      * Describe has varname, vartype, varlabel and dummy for isnumeric
      * this forms the base of the codebook, for which we will add:
      * - count and distinct
      * - summary stats for numerica variables
      describe, replace clear
      tempfile codebook_base
      save `codebook_base'
      
      * The vnumlist has the original position of the filtered position
      * of all the numeric variables, will be handy later
      keep if isnumeric
      forvalue i = 1/`=_N' {
         local vnumlist "`vnumlist' v`=position[`i']'"
      }
      keep position
      gen vnum = _n
      tempfile cheatsheet
      save `cheatsheet'
    
    * Back to the original dataset
    restore
    
    * Rename each variable according to its position ie: v1 v2 ... vI
    * while storing the original varname as a local
    local i=1
    ds
    foreach v in `r(varlist)' {
      rename `v' v`i'
      local varnamev`i' "`v'"
      local ++i
    }
    * Total number of variables stored in i has passed over
    local --i
    
    * Count of nonmissing values and distinct values for each variable
    * (including non-numeric variables), stored as locals
    foreach v of varlist * {
      tempvar touse vals
      mark `touse'
      markout `touse' `v', strok 
      bys `touse' `v' : gen byte `vals' = (_n == 1)
      sum `vals' if `touse', meanonly
      local count`v' = r(N)
      local distinct`v' = r(sum)
    }
    
    * Collapse the numeric variables for each of the stats to include
    * storing results as tempfiles
    local stats "mean min max"
    tempfile `stats'
    foreach s in `stats' {
      preserve
        collapse (`s') `vnumlist', fast
        save ``s''
      restore
    }        
    
    * Preserve the original dataset, but with the varnames changed to v1...vI
    preserve
    
      * Manipulate numeric variable stats into a single tempfile
      use `mean', clear
      append using `min'
      append using `max'
      xpose, clear
      rename (v1 v2 v3) (mean min max)
      gen vnum=_n    
      tempfile stats
      save `stats'
      
      * Back to the base of the codebook
      use `codebook_base', clear
      
      * Placeholders for the count of non-missing and distinct values
      gen long count    = .
      gen long distinct = .
      * Substitute for the actual values stored as locals
      forvalues j = 1/`i' {
        replace count    = `countv`j''    if _n == `j'
        replace distinct = `distinctv`j'' if _n == `j'
      }
      
      * Before bringing in the numeric variables stats, must know how they were
      * ordered (vnum), which was stored in the cheatsheet
      merge 1:1 position using `cheatsheet', nogen
      * With vnum we can attach the correct stats to each of those variables
      merge m:1 vnum using `stats', nogen
      
      * Organize the codebook file
      sort position
      drop vallab vnum position format
      rename (name varlab) (varname varlabel)
      order varname type isnumeric varlabel count-max
      
      label var count "nonmissing observations"
      label var distinct "distinct values"
      label var mean "mean"
      label var min "minimum"
      label var max "maximum"
      
      * Display the codebook that will be saved
      noi disp as text _n "Codebook"
      noi list, noobs clean table
      noi disp as text _n 
      
      * Save codebook file, acording to the specified format
      if "`r_extension'" == ".dta" noi save `"`using'"', `replace'
      if "`r_extension'" == ".csv" noi export delimited `"`using'"', `replace'
      if inlist("`r_extension'", ".xlsx", ".xls") noi export excel `"`using'"', sheet("codebook") `replace' firstrow(variables)
    
    * Back to the original data, but with the varnames changed to v1...vI
    restore
    
    * Revert back to original varnames
    forvalues j = 1/`i' {
      rename v`j' `varnamev`j''
    }
    
    * Now we have exactly the same dataset as in the beginning
    
  }
  
end

* Not terribly elegant, but does the trick
cap mata: mata drop split_using()

*------------------------------- MATA -----------------------------------------

mata:
mata set matastrict on

void split_using(string scalar using2split) {
// using2split broken into macros: r(path), r(filename), r(extension)
// if the path is not absolute, it is autocompleted with pwd

  string scalar path,
                filename,
                extension
  pragma unset path
  pragma unset filename
  pragma unset extension
  
  // Autocomplete to absolute path if relative path
  if (pathisabs(using2split) == 0) {
    using2split = pathjoin(pwd(), using2split)
  }

  // Attempt to extract file extension
  extension = pathsuffix(using2split)

  // If there is no extension, the whole thing is a path (filename is empty)
  if (extension == "") {
    filename = ""
    path = using2split
  }
  // Else, split the path and filename
  else {
    pathsplit(using2split, path, filename)
  }
  
  st_rclear()
  st_global("r(folder)"   , path)
  st_global("r(filename)" , filename)
  st_global("r(extension)", extension)
  
}
end