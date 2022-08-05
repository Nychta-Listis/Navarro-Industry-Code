using JuMP
import XLSX

# Solvers I have tested and managed to get to work: Ipopt, EAGO (slow)
solver = "Ipopt"
eval(Meta.parse("import $solver"))

##################
# PREPARING DATA #
##################


# Work Path
#egridsPath = "C:\\Users\\anavarro\\Box\\NaturalGas\\Model 3\\egrid2020_data_metric.xlsx"
# Personal Comp Path
egridsPath = "/Users/alex/Library/CloudStorage/Box-Box/NaturalGas/Model 3/egrid2020_data_metric.xlsx"




# Types of Fuels being considered
global PLNTFuel = ["DFO","JF","KER","LFG","MSW","NG","OBL","PC","RFO","WDS","WO","BIT", "COG", "LIG", "RC", "SUB", "WC"]
# Column labels to be extrated from file (keep PLGENA* in this order and consecutive)
global needPLNT = ["ORISPL", "SECTOR", "PLPRMFL", "PLFUELCT", "CAPFAC", "NAMEPCAP","UNCO2",  "PLHTRT","PLGENACL","PLGENAOL","PLGENAGS","PLGENABM","PLGENAOF","PLGENAOP"]
#global needPLNT = ["ORISPL", "SECTOR", "PLPRMFL", "PLFUELCT", "UNCO2", "PLHTRT","PLGENACL","PLGENAOL","PLGENAGS"] # UNCO2 <- PLCO2AN
NumPara = length(findall(x -> occursin(r"PLGENA*", x), needPLNT))
# Internal Data Dict (Simply returns column of entry in internal data Array)
global intPLNT20Dict = Dict()
for i in 1:length(needPLNT)
    intPLNT20Dict[needPLNT[i]] = i
end
# Row number with desired column labels
columnLabelRow = 2
# Ignore all rows with missing data 
ignoreMissing = true # (Default = true)
PLNTDict = Dict()

# Opens data file and stores Plant information in relevent data structures
# PLNT20 contains all Plants that are used, and can just be interated by rows [Data Column, Plant]
# PLNTDict Dictionary containing all plants with keys as the ORIS ID
PLNTORIS = Set()
XLSX.openxlsx(egridsPath, enable_cache=false) do f
    sheet = f["PLNT20"]
    cellType = 0
    global PLNT20Dict = Dict()
    global PLNTSectDict = Dict()
    firstValidMark = true
    n=1
    if ignoreMissing # Two cases of if ignoreMissing, should be identical other than missing conditions
        for r ∈ XLSX.eachrow(sheet)
            # Conditions after "true &&" represent filters ("true &&" can be removed for speedup, filters changed depending on data of interest)
            if n>columnLabelRow && !firstValidMark && (0==sum(map(x -> x == Missing, typeof.([r[PLNT20Dict[i]] for i ∈ needPLNT])))) && true && (r[PLNT20Dict["SECTOR"]] == "Electric Utility") && (r[PLNT20Dict["PLPRMFL"]] ∈ PLNTFuel) && (r[PLNT20Dict["CAPFAC"]] != 0)
                PLNT20 = cat(PLNT20, [r[PLNT20Dict[i]] for i ∈ needPLNT],dims=2)
                PLNTDict[r[PLNT20Dict["ORISPL"]]] = [r[PLNT20Dict[i]] for i ∈ needPLNT]
                push!(PLNTORIS, r[PLNT20Dict["ORISPL"]])
            elseif firstValidMark && n>columnLabelRow && (0==sum(map(x -> x == Missing, typeof.([r[PLNT20Dict[i]] for i ∈ needPLNT])))) && true && (r[PLNT20Dict["SECTOR"]] == "Electric Utility") && (r[PLNT20Dict["PLPRMFL"]] ∈ PLNTFuel) && (r[PLNT20Dict["CAPFAC"]] != 0)
                global PLNT20 = [r[PLNT20Dict[i]] for i ∈ needPLNT]
                PLNTDict[r[PLNT20Dict["ORISPL"]]] = [r[PLNT20Dict[i]] for i ∈ needPLNT]
                push!(PLNTORIS, r[PLNT20Dict["ORISPL"]])
                firstValidMark = false
            elseif n == columnLabelRow # Create dictionary for taking in column labels and returning the column number
                m=1
                while !(r[m] isa Missing)
                    PLNT20Dict[r[m]] = m
                    m+=1
                end
            end # Create dictionary for that returns sector when provided with ORIS ID (Used for when working with generators)
            if n>columnLabelRow
                if typeof(r[PLNT20Dict["SECTOR"]]) == Missing
                    PLNTSectDict[r[PLNT20Dict["ORISPL"]]] = "NULL"
                else
                    PLNTSectDict[r[PLNT20Dict["ORISPL"]]] = r[PLNT20Dict["SECTOR"]]
                end
            end
            n += 1
        end
    else # Two cases of if ignoreMissing, should be identical other than missing conditions
        for r in XLSX.eachrow(sheet)
            # Conditions after "true &&" represent filters ("true &&" can be removed for speedup, filters changed depending on data of interest)
            if n>columnLabelRow && !firstValidMark && true && (r[PLNT20Dict["SECTOR"]] == "Electric Utility") && (r[PLNT20Dict["PLPRMFL"]] ∈ PLNTFuel) && (r[PLNT20Dict["CAPFAC"]] != 0)
                PLNT20 = cat(PLNT20, [r[PLNT20Dict[i]] for i ∈ needPLNT],dims=2)
                push!(PLNTORIS, r[PLNT20Dict["ORISPL"]])
            elseif firstValidMark && n>columnLabelRow && true && (r[PLNT20Dict["SECTOR"]] == "Electric Utility") && (r[PLNT20Dict["PLPRMFL"]] ∈ PLNTFuel) && (r[PLNT20Dict["CAPFAC"]] != 0)
                global PLNT20 = [r[PLNT20Dict[i]] for i ∈ needPLNT]
                push!(PLNTORIS, r[PLNT20Dict["ORISPL"]])
                firstValidMark = false
            elseif n == columnLabelRow
                m=1
                while !(r[m] isa Missing) # Create dictionary for taking in column labels and returning the column number
                    PLNT20Dict[r[m]] = m
                    m+=1
                end
            end
            if n>columnLabelRow
                if typeof(r[PLNT20Dict["ORISPL"]]) == Missing
                    PLNTSectDict[r[PLNT20Dict["ORISPL"]]] = "NULL"
                else
                    PLNTSectDict[r[PLNT20Dict["ORISPL"]]] = r[PLNT20Dict["SECTOR"]]
                end
            end
            n += 1
        end
    end
end

println(sum(map(x -> x == Missing, typeof.(PLNT20))))
#println(typeof.(PLNT20[1:5,1:192]))
println(size(PLNT20))

# Dictionary for taking in disaggregated fuel types and returning the category to which they have been assigned
PLNT20FuelDict = Dict(
    "AB"  => "BM", "BLQ" => "BM", "LFG" => "BM",
    "MSW" => "BM", "OBG" => "BM", "OBL" => "BM",
    "OBS" => "BM", "SLW" => "BM", "WDL" => "BM",
    "WDS" => "BM",
    "NG"  => "NG", "PG"  => "NG",
    "DFO" => "PT", "JF"  => "PT", "KER" => "PT",
    "PC"  => "PT", "RFO" => "PT", "SGP" => "PT",
    "TDF" => "PT", "WO"  => "PT",
    "BIT" => "CL", "COG" => "CL", "LIG" => "CL",
    "RC"  => "CL", "SUB" => "CL", "WC" => "CL"
)
PLNTFCAT = ["BM","NG","PT", "CL"]
for i ∈ 1:size(PLNT20,2)
    PLNT20[intPLNT20Dict["PLPRMFL"],i] = PLNT20FuelDict[PLNT20[intPLNT20Dict["PLPRMFL"],i]]
end

for (key, value) ∈ PLNTDict
    value[intPLNT20Dict["PLPRMFL"]] = PLNT20FuelDict[value[intPLNT20Dict["PLPRMFL"]]]
end
###############
# CO2 FACTORS #
###############

println("\n\nIdentifying CO2 Factors")

co2fmodel = eval(Meta.parse("Model($solver.Optimizer)"))
#1 - coal, 2 - oil, 3 - gas,  4 - biomass, 5 - other fossil, 6 - unknown purchased
@variable(co2fmodel, 0 ≤ CFS[1:NumPara,PLNTFCAT]) # CFS: kg/mmBTU
#@variable(model, CFI) # Assume nonzero residual on linear fit
CFI = 0 # (metric tons) # Assume zero residual on linear fit

# Fitted CO2 Generated (metric tons)    kg/mmBTU         MWh                             BTU/kWH      1 mmBTU/1000000 BTU    (1000 kWH/MWh    t/1000kg)
CO2VecCalc = @expression(co2fmodel, [sum([CFS[j,PLNT20[intPLNT20Dict["PLPRMFL"],i]]*PLNT20[j+intPLNT20Dict["PLGENACL"]-1, i]*PLNT20[intPLNT20Dict["PLHTRT"], i] for j ∈ 1:NumPara])/1000000+CFI for i ∈ 1:size(PLNT20,2)])
# Minimize error in CO2 Generation
# (Sum of 0.01*CFS term added so that irrelevant factors will be minimized to 0 (Tested, and terms that matter are relatively unaffected))
@objective(co2fmodel, Min, sum((CO2VecCalc .- PLNT20[intPLNT20Dict["UNCO2"],1:end]).^2) + sum(0.01*[CFS[i,j] for i ∈ 1:size(CFS,1), j ∈ PLNTFCAT]))

optimize!(co2fmodel)

println(solution_summary(co2fmodel))

# Labels for the sake of printing
lst = ["Coal", "Oil", "Gas", "Biomass", "Other Fossil", "Unknown"]

ybar = 1/size(PLNT20,2)*sum(PLNT20[intPLNT20Dict["UNCO2"],1:end])
Rsqr = 1 - sum(([sum([value(CFS[j,PLNT20[intPLNT20Dict["PLPRMFL"],i]])*PLNT20[j+intPLNT20Dict["PLGENACL"]-1, i]*PLNT20[intPLNT20Dict["PLHTRT"], i] for j ∈ 1:NumPara])/1000000+value(CFI) for i ∈ 1:size(PLNT20,2)] .- PLNT20[intPLNT20Dict["UNCO2"],1:end]).^2)/sum((ybar .- PLNT20[intPLNT20Dict["UNCO2"],1:end]).^2)
println("R² = $Rsqr")


# Display outputs: For each output the plant level information assumes multiple different fuel inputs for a single plant,
# and so, CO2 factors are calculated for each of these. If single values are desired, then just take the corresponding column,
# for example, for Coal, look at the 4th row 1st column, and take that as your value as dependent on that one fuel source
CFSArray = [((0.5 ≤ value(CFS[i,j]) ≤ 10000) ? value(CFS[i,j]) : 0) for i∈1:size(CFS,1), j∈PLNTFCAT]
println(lst)
for j ∈ 1:length(PLNTFCAT)
    string = "$(PLNTFCAT[j]):"
    for i ∈ 1:size(CFSArray,1)
        string *= " $(CFSArray[i,j])"
    end
    println(string)
end


####################
# CAPACITY FACTORS #
####################


# Types of Generator Statuses being considered
global GENStat = ["OP", "SB"]
# Types of Fuels being considered
global GENFuel = ["AB","BLQ","LFG","MSW","OBG","OBL","OBS","SLW","WDL","WDS","NG","PG","DFO","JF","KER","PC","RFO","SGP","TDF","WO","BIT","COG","LIG","RC","SUB","WC"]
# Column labels to be extrated from file
global needGEN = ["ORISPL", "GENSTAT", "PRMVR", "FUELG1", "NAMEPCAP", "CFACT", "GENNTAN"]
# Internal Data Dict (Simply returns column of entry in internal data Array)
global intGEN20Dict = Dict()
for i ∈ 1:length(needGEN)
    intGEN20Dict[needGEN[i]] = i
end
# Row number with desired column labels
columnLabelRow = 2
# Ignore all rows with missing data 
ignoreMissing = true # (Default = true)
# Dictionary used for estimating the disaggregated average capacity factors, [sum of capacity factors multiplied by energy output, sum of energy outputs]
Disagg = Dict()

XLSX.openxlsx(egridsPath, enable_cache=false) do f
    sheet = f["GEN20"]
    cellType = 0
    global GEN20Dict = Dict()
    firstValidMark = true
    n=1
    if ignoreMissing # Two cases of if ignoreMissing, should be identical other than missing conditions
        for r in XLSX.eachrow(sheet)
#            if n>5
#                println(r[GEN20Dict["ORISPL"]], (PLNTSectDict[r[GEN20Dict["ORISPL"]]] == "Electric Utility") && (r[GEN20Dict["GENSTAT"]] ∈ GENStat))
#            end
            # Conditions after "true &&" represent filters ("true &&" can be removed for speedup, filters changed depending on data of interest)
            if n>columnLabelRow && !firstValidMark && (0==sum(map(x -> x == Missing, typeof.([r[GEN20Dict[i]] for i ∈ needGEN])))) && true && (r[GEN20Dict["FUELG1"]] ∈ GENFuel) && (PLNTSectDict[r[GEN20Dict["ORISPL"]]] == "Electric Utility") && (r[GEN20Dict["GENSTAT"]] ∈ GENStat) && (r[GEN20Dict["PRMVR"]] != "CE") && (r[GEN20Dict["GENNTAN"]] > 0)
                GEN20 = cat(GEN20, [r[GEN20Dict[i]] for i ∈ needGEN],dims=2)
                if [r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]] in keys(Disagg)
                    Disagg[[r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]]] += [r[GEN20Dict["GENNTAN"]]^2/r[GEN20Dict["NAMEPCAP"]],r[GEN20Dict["GENNTAN"]]]
                else
                    Disagg[[r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]]] = [r[GEN20Dict["GENNTAN"]]^2/r[GEN20Dict["NAMEPCAP"]],r[GEN20Dict["GENNTAN"]]]
                end
            elseif firstValidMark && n>columnLabelRow && (0==sum(map(x -> x == Missing, typeof.([r[GEN20Dict[i]] for i ∈ needGEN])))) && true && (r[GEN20Dict["FUELG1"]] ∈ GENFuel) && (PLNTSectDict[r[GEN20Dict["ORISPL"]]] == "Electric Utility") && (r[GEN20Dict["GENSTAT"]] ∈ GENStat) && (r[GEN20Dict["PRMVR"]] != "CE") && (r[GEN20Dict["GENNTAN"]] > 0)
                global GEN20 = [r[GEN20Dict[i]] for i ∈ needGEN]
                if [r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]] in keys(Disagg)
                    Disagg[[r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]]] += [r[GEN20Dict["GENNTAN"]]^2/r[GEN20Dict["NAMEPCAP"]],r[GEN20Dict["GENNTAN"]]]
                    
                else

                    Disagg[[r[GEN20Dict["PRMVR"]],r[GEN20Dict["FUELG1"]]]] = [r[GEN20Dict["GENNTAN"]]^2/r[GEN20Dict["NAMEPCAP"]],r[GEN20Dict["GENNTAN"]]]
                end
                firstValidMark = false
            elseif n == columnLabelRow # Create dictionary for taking in column labels and returning the column number
                m=1
                while !(r[m] isa Missing)
                    GEN20Dict[r[m]] = m
                    m+=1
                end
            end
            n += 1
        end
    else # Two cases of if ignoreMissing, should be identical other than missing conditions
        for r in XLSX.eachrow(sheet)
            # Conditions after "true &&" represent filters ("true &&" can be removed for speedup, filters changed depending on data of interest)
            if n>columnLabelRow && !firstValidMark && true && (r[GEN20Dict["FUELG1"]] ∈ GENFuel) && (PLNTSectDict[r[GEN20Dict["ORISPL"]]] == "Electric Utility") && (r[GEN20Dict["GENSTAT"]] ∈ GENStat) && (r[GEN20Dict["PRMVR"]] != "CE")
                GEN20 = cat(GEN20, [r[GEN20Dict[i]] for i ∈ needGEN],dims=2)
            elseif firstValidMark && n>columnLabelRow && true && (r[GEN20Dict["FUELG1"]] ∈ GENFuel) && (PLNTSectDict[r[GEN20Dict["ORISPL"]]] == "Electric Utility") && (r[GEN20Dict["GENSTAT"]] ∈ GENStat) && (r[GEN20Dict["PRMVR"]] != "CE")
                global GEN20 = [r[GEN20Dict[i]] for i ∈ needGEN]
                firstValidMark = false
            elseif n == columnLabelRow # Create dictionary for taking in column labels and returning the column number
                m=1
                while !(r[m] isa Missing)
                    GEN20Dict[r[m]] = m
                    m+=1
                end
            end
            n += 1
        end
    end
end
# Dictionary takes in mover and returns the aggregated mover category to which it belongs
# Seperated Gas and Steam Turbine 
#=
GEN20MoverDict = Dict(
    "CA" => "CC", "CC" => "CC", "CS" => "CC", "CT" => "CC",
    "FC" => "GT", "GT" => "GT", "IC" => "GT", "OT" => "GT",
    "ST" => "ST"
)
=#

# Dictionary takes in mover and returns the aggregated mover category to which it belongs
# Combined Gas and Steam Turbine
GEN20MoverDict = Dict(
    "CA" => "CC", "CC" => "CC", "CS" => "CC", "CT" => "CC",
    "FC" => "GT", "GT" => "GT", "IC" => "GT", "OT" => "GT",
    "ST" => "GT" # Steam turbine has been added to gas turbines.
)

# Dictionary takes in fuel and returns the aggregated fuel category to which it belongs
GEN20FuelDict = Dict(
    "AB"  => "BM", "BLQ" => "BM", "LFG" => "BM",
    "MSW" => "BM", "OBG" => "BM", "OBL" => "BM",
    "OBS" => "BM", "SLW" => "BM", "WDL" => "BM",
    "WDS" => "BM",
    "NG"  => "NG", "PG"  => "NG",
    "DFO" => "PT", "JF"  => "PT", "KER" => "PT",
    "PC"  => "PT", "RFO" => "PT", "SGP" => "PT",
    "TDF" => "PT", "WO"  => "PT",
    "BIT" => "CL", "COG" => "CL", "LIG" => "CL",
    "RC"  => "CL", "SUB" => "CL", "WC" => "CL"
)

# Dictionary takes in fuel and mover and returns the category to which it is ultimately assigned
GEN20LabelDict = Dict(
    ["BM","CC"] => "BM", ["BM","GT"] => "BM", ["BM","ST"] => "BM",
    ["NG","CC"] => "NGCC", ["NG","GT"] => "NGGT", ["NG","ST"] => "NGST",
    ["PT","CC"] => "PT", ["PT","GT"] => "PT", ["PT","ST"] => "PT",
    ["CL","CC"] => "CL", ["CL","GT"] => "CL", ["CL","ST"] => "CL"
)

# Combined mover and fuel labels being modeled
GEN20Labels = ["BM", "NGCC", "NGGT", "PT", "CL"]

# GEN20Labels = ["BM", "NGCC", "NGGT", "NGST", "PT", "CL"]
for i ∈ 1:size(GEN20,2)
    GEN20[intGEN20Dict["PRMVR"],i] = GEN20MoverDict[GEN20[intGEN20Dict["PRMVR"],i]]
    GEN20[intGEN20Dict["FUELG1"],i] = GEN20FuelDict[GEN20[intGEN20Dict["FUELG1"],i]]
    #println(GEN20[1:end,i])
end

# Mover and Fuel Labels
MDMVR = ["CC","GT","ST"]
MDFUE = ["BM","NG","PT", "CL"]

#=
println("\n\nIdentifying Capacity Factors")

fuelVec = String[]
for i ∈ GEN20[intGEN20Dict["PLFUELCT"],1:end]
    i in fuelVec ? continue : push!(fuelVec,i)
end
println(fuelVec)
=#

totalPower = sum([GEN20[intGEN20Dict["GENNTAN"], i] for i ∈ 1:size(GEN20,2)])
println(totalPower)
cfmodel = eval(Meta.parse("Model($solver.Optimizer)"))
@variable(cfmodel, 0 ≤ CF[GEN20Labels] ≤ 1) # CF

# ELE GEN (MWh)                                                             MW                              dy/yr hr/dy
POWVec = @expression(cfmodel, [CF[GEN20LabelDict[[GEN20[intGEN20Dict["FUELG1"],i],GEN20[intGEN20Dict["PRMVR"],i]]]] * GEN20[intGEN20Dict["NAMEPCAP"],i] * 365 * 24 for i ∈ 1:size(GEN20,2)])


# Minimize error in Power Generation
#@constraint(cfmodel, totalPower == sum(POWVec))
# Possible test for reasonable validity, try removing above condition and see how elec err is affected, as then problem is simply regression.
@objective(cfmodel, Min, sum((POWVec .- [GEN20[intGEN20Dict["GENNTAN"],i] for i ∈ 1:size(GEN20,2)]).^2))

optimize!(cfmodel)

println(solution_summary(cfmodel))

# Print Capacity factors
for i ∈ GEN20Labels
    println("$i = $(value(CF[i]))")
end

println("\nElec. Err. = $((sum(value.(POWVec))-totalPower)/totalPower)")

# Code for estimating CO2 emissions from all generators considered
# NOTE: NOT NECESSARILY SECTOR WIDE, AS SOME GENERATORS ARE EXCLUDED BECAUSE OF MISSING ENTRIES
#=
global CO2Vec = []
for i ∈ 1:size(GEN20,2)
    if GEN20[intGEN20Dict["ORISPL"],i] ∈ PLNTORIS
        push!(CO2Vec,GEN20[intGEN20Dict["NAMEPCAP"],i]*value(CF[GEN20LabelDict[[GEN20[intGEN20Dict["FUELG1"],i],GEN20[intGEN20Dict["PRMVR"],i]]]]) * PLNTDict[GEN20[intGEN20Dict["ORISPL"],i]][intPLNT20Dict["PLHTRT"]] *365*24* sum([value(CFS[j,PLNTDict[GEN20[intGEN20Dict["ORISPL"],i]][intPLNT20Dict["PLPRMFL"]]])*PLNTDict[GEN20[intGEN20Dict["ORISPL"],i]][j+intPLNT20Dict["PLGENACL"]-1] for j ∈ 1:NumPara])/sum([PLNTDict[GEN20[intGEN20Dict["ORISPL"],i]][j+intPLNT20Dict["PLGENACL"]-1] for j ∈ 1:NumPara])/1000000)
#    else
#        println(GEN20[intGEN20Dict["ORISPL"],i])
    end
end
=#


# Weighted average of disaggregated sets is obtained by dividing the first entry by the second for the Disagg dict (And appropriate conversion factors) (For comparison to aggregated values)
for (key, value) ∈ Disagg
    println("$key:  $(value[1]/(365*24*value[2]))")
end