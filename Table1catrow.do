capture program drop table1catrow
program define table1catrow
	syntax varlist [if], by(varname) [replace] 
    //use the "putexcel" command

    //setup
    //putexcel will write to table1.xlsx
    //worksheet table1
    //"modify" = keep the existing Excel file if it exists
    //           (just overwrite one cell at a time)
    if "`if'"=="" {
		putexcel set "table1_by`by'.xlsx", sheet("table1_row") modify
	}
	else {
		putexcel set "table1_by`by'_`if'.xlsx", sheet("table1_row") modify
	}	
	
	putexcel A1="Table 1. Row percentages of categorical variables by `by' `if'", bold
	putexcel B1="`by'"
	putexcel (B1:C1), merge hcenter vcenter
	
	//This needs to be 3 instead of 1 because the first row has the title, and the second row has labels of the by variable
	local rownum=3
	
foreach v of varlist `varlist' {
	if "`if'"=="" {
		tab `v' if !missing(`by'), matcell(rowtotals)
		tab `by' if !missing(`v'), matcell(coltotals)
		tab `v' `by', matcell(cellcounts)
		local rowcount=r(r)
		local colcount=r(c)
	}
	else {
		tab `v' `if' & !missing(`by'), matcell(rowtotals)
		tab `by' `if' & !missing(`v'), matcell(coltotals)
		tab `v' `by' `if', matcell(cellcounts)
		local rowcount=r(r)
		local colcount=r(c)
	}	
		
		//Add variable label:
		local RowVarLabel : var label `v'
		local Cell = char(64 + 1) + string(`rownum')
		putexcel `Cell' = "`RowVarLabel'", right
	
		
		local RowValueLabel : value label `v'
		levelsof `v', local(RowLevels)
		
		local ColValueLabel : value label `by'
		levelsof `by', local(ColLevels)
		
		forvalues row=1/`rowcount' {

		//Add level labels:
			local RowValueLabelNum = word("`RowLevels'", `row')
			local CellContents : label `RowValueLabel' `RowValueLabelNum'
			local Cell = char(64 + 1) + string(`rownum'+1)
			putexcel `Cell' = "`CellContents'", right
		 
		 //Add row totals:
			local CellContents = rowtotals[`row',1]
			local Cell = char(64 + `colcount' + 2) + string(`rownum' + 1)
			putexcel `Cell' = "`CellContents'", hcenter
		 
		
			forvalues col=1/`colcount' {
		
				//Add counts and percentages to each cell:
					local cellcount = cellcounts[`row',`col']
					local cellpercent = string(100*`cellcount'/rowtotals[`row',1],"%9.1f")
					local CellContents = "`cellcount' (`cellpercent'%)"
					local Cell = char(64 + `col' + 1) + string(`rownum' + 1)
					putexcel `Cell' = "`CellContents'", right

				//Add column value labels:
					local ColValueLabelNum = word("`ColLevels'", `col')
					local CellContents : label `ColValueLabel' `ColValueLabelNum'
					local Cell = char(64 + `col' + 1) + string(2)
					putexcel `Cell' = "`CellContents'", hcenter
						
				//Also add a column header for "Total"
					local Cell = char(64 + `col' + 2) + string(2)
					putexcel `Cell' = "Total", hcenter
									
				//Add total ns to each column at the bottom of the table
				//(This will produce n's at the last row of each variable)
				if `row'==`rowcount' {		 
					local CellContents = coltotals[`col',1]
					local Cell = char(64 + `col' + 1) + string(`rownum' + 2)
					putexcel `Cell' = "`CellContents'", hcenter
				}
			}
		local rownum=`rownum'+1	
		}
		local rownum=`rownum'+2
		
	}	
end

	/*
	table1catrow new_race final_sex, by(screener_own_wp) replace
	
	
	
	
