import XLSX
import DataFrames
using Flux, Plots
using DataAPI
using BSON: @load
using BSON: @save
using Flux: train!
import Base

#=
TODO:

    Comment Functions
        Description
        inputs
        Outputs
    Comment internal Behavior of Functions
    Create several Example Programs

    Create function for running data on network
        (Takes in only input data and returns model prediction)
    Create function for not training or running, but rather reporting performance of network on data
        (Takes in input and output data and returns stuff like error)

Possible:

    Go through .xlsx files and remove all but the used columns (preserve the files with all just in case)
        Would need to rewrite inputData function to have mode for either full or reduced file
        Probably a necessary change in order to be usable with an actual model
            since we don't want to have to deal with a bunch of unused columns
=#

#= ACCEPTED ASM FILE FORMAT
File Format: .xlsx

Rows
Row 1: LABELS
All other rows: DATA

- denotes default network output
= denotes default network input
@ Used for data input ()

Columns
#   A:  Geographic area name
#@  B:  NAICS code
#   C:  Meaning of NAICS code
1   D:  Year
2=  E:  Number of employees
3=  F:  Annual payroll ($1,000)
4   G:  Total fringe benefits ($1,000)
5   H:  Employer's cost for health insurance ($1,000)
6   I:  Employer's cost for defined benefit pension plans ($1,000)
7   J:  Employer's cost for defined contribution plans ($1,000)
8   K:  Employer's cost for other fringe benefits ($1,000)
9=  L:  Production workers avg per year
10= M:  Production workers hours (1,000)
11  N:  Production workers wages ($1,000)
12= O:  Total cost of materials ($1,000)
13= P:  Materials, parts, containers, packaging, etc. used ($1,000)
14  Q:  Cost of resales ($1,000)
15  R:  Cost of purchased fuels ($1,000)
16= S:  Purchased electricity ($1,000)
17  T:  Contract Work ($1,000)
18= U:  Quantity of electricity purchased for heat and power (1,000 kWh)
19= V:  Quantity of generated electricity (1,000 kWh)
20= W:  Quantity of electricity sold or transferred (1,000 kWh)
21  X:  Value added ($1,000)
22= Y:  Total inventories, beginning of year ($1,000)
23  Z:  Finished goods inventories, beginning of year ($1,000)
24  AA: Work-in-process inventories, beginning of year ($1,000)
25  AB: Materials and supplies inventories, beginning of year ($1,000)
26= AC: Total inventories, end of year ($1,000)
27  AD: Finished goods inventories, end of year ($1,000)
28  AE: Work-in-process inventories, end-of-year  ($1,000)
29  AF: Materials and supplies inventories, end of year ($1,000)
30= AG: Total capital expenditures (new and used) ($1,000)
31  AH: Capital expenditures on buildings and other structures (new and used) ($1,000)
32  AI: Capital expenditures on machinery and equipment (new and used) ($1,000)
33  AJ: Capital expenditures on automobiles, trucks, etc. for highway use ($1,000)
34  AK: Capital expenditures on computers and peripheral data processing equipment ($1,000)
35  AL: Capital expenditures on all other machinery and equipment ($1,000)
36= AM: Total rental payments ($1,000)
37  AN: Rental payments for buildings and other structures ($1,000)
38= AO: Temporary staff and leased employee expenses ($1,000)
39  AP: Expensed computer hardware and other equipment ($1,000)
40  AQ: Expensed purchases of software ($1,000)
41  AR: Data processing and other purchased computer services ($1,000)
42  AS: Communication services ($1,000)
43  AT: Repair and maintenance services of buildings and/or machinery ($1,000)
44  AU: Refuse removal (including hazardous waste) services ($1,000)
45  AV: Advertising and promotional services ($1,000)
46  AW: Purchased professional and technical services ($1,000)
47= AX: Taxes and license fees ($1,000)
48= AY: All other expenses ($1,000)
49  AZ: Rental payments for machinery and equipment ($1,000)
50  BA: Total other expenses ($1,000)
51- BB: Total value of shipments ($1,000)

=#


#= importData
Basic function for taking in ASM datafiles from a single directory
and preparing it for use by the Neural Network for either training or simply as a value

inputs:
path: STRING: Path to directory containing either training or input data
mode: INTEGER: 0 to denote training data or 1 to denote testing data or 2 to denote custom label
NAICS: STRING: NAICS group to be taken as input, can be any level. NAICS Sectors will be taken in with same length or longer that fall within it
            eg. if "325" is used, then "325" and "3251" and up through all subgroups within each of those, such as "325188"
inVec: VECTOR[INTEGER]: Columns from datafiles to be used as network input (Default described above in file format for full file)
outVec: VECTOR[INTEGER]: Columns from datafiles to be used as network output (Default described above in file format for full file)
inlabel: STRING: Optional custom label for identifying input files (mode must be set to 2 if this option is used) (Default "IN_ASM")

TO ADD: reduced: BOOL: Is a reduced data file being used (Default false)

NOTE: Input files are indentified by the first part of the file name.
If the file name starts with "TRAIN_ASM" it is recognized as training data.
If the file name starts with "TEST_ASM" it is recognized as TEST data.
If mode is set to 2, the file name starts with the value of inlabel

returns:
data:

indat:

outdat:

=#

function importData(path,mode,NAICS;inVec = [2,3,9,10,12,13,16,18,19,20,22,26,30,36,38,47,48],outVec = [51], inlabel = "IN_ASM")
    if mode == 0 # train
        label = "TRAIN_ASM"
    elseif mode == 1 # test
        label = "TEST_ASM"
    elseif mode == 2 # custom label
        label = inlabel
    else
        throw(ErrorException(label, "Mode Selected Does Not Exist"))
    end
    cd(path)
    ASMPaths = []
    for file ∈ readdir()
        if length(file) ≥ 10 && file[1:length(label)] == label
            push!(ASMPaths, file)
        end
    end
    global dataVec = []
    Columns = [i for i ∈ 4:54]
    colDict = Dict()
    for i in 1:length(Columns)
        colDict[Columns[i]] = i
    end
    for file ∈ ASMPaths
        XLSX.openxlsx(file, enable_cache=false) do f
            sheet = f["Data"]
            n=1
            for r ∈ XLSX.eachrow(sheet)
                if n == 1
                    global labels = [(r[i]) for i ∈ Columns]
                elseif dataVec == [] && n>1 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICS) && "$(r["B"])"[1:length(NAICS)] == NAICS
                    global dataVec = [formatData(r[i]) for i ∈ Columns]
                elseif n>=2 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICS) && "$(r["B"])"[1:length(NAICS)] == NAICS
                    dataVec = cat(dataVec,[formatData(r[i]) for i ∈ Columns],dims = 2)
                end
                n += 1
            end
        end
    end
    dataVec = map(x -> typeof(x) <: Number ? x : 0, dataVec)
    dataVec = [i == 1 ? dataVec[i,j]-2000  : log(dataVec[i,j]+1) for i ∈ 1:size(dataVec,1), j ∈ 1:size(dataVec,2)]
    dataF = DataFrames.DataFrame(dataVec', labels)
    inDataF = DataFrames.select(dataF,inVec)
    outDataF = DataFrames.select(dataF,outVec)
    insize = size(inDataF,2)
    outsize = size(outDataF,2)
    indat = [[inDataF[i, j] for j ∈ 1:insize] for i ∈ 1:size(dataVec,2)]
    outdat = [[outDataF[i, j] for j ∈ 1:outsize] for i ∈ 1:size(dataVec,2)]
    data = [(indat, outdat)]
    return (data, indat, outdat)
end

function createModel(insi, outsi)
    #slayer = Flux.Scale(ones(insi),false,relu)
    layerIn = Flux.Dense(insi,outsi,relu) # ; init = Flux.rand32
    #layerOut = Flux.Dense(2,outsi,relu)
    model = Flux.Chain(layerIn)
    return model
end

function formatData(x)
    try
        return Meta.parse(replace(x, "," => "_"))
    catch
        return x
    end
end

function trainepochs(modelV, epochs, data; label = "", save = true, printstat = true,modelName = "model_0")
    # How often (in epochs) the model creates a checkpoint and modifies training rate
    checklen = 1000
    modelVe = deepcopy(modelV)
    intrain = data[1][1]
    outtrain = data[1][2]
    if save; @save "$(modelName)_b.bson" modelVe; end
    if printstat
        for epoch ∈ 0:checklen:(epochs-checklen)
            prevLoss = Base.invokelatest(modelVe[2],intrain,outtrain)
            for epoc ∈ 1:checklen
                train!(modelVe[2], Flux.params(modelVe[1][1:end]), data, modelVe[6])
            end
            println("Epoch $(epoch+checklen): $(Base.invokelatest(modelVe[2],intrain,outtrain))")
            if Base.invokelatest(modelVe[2],intrain,outtrain) > prevLoss
                println("Model performance worsened, resetting to previous point")
                @load "$(modelName)_b.bson" modelVe
                modelVe[3] /= ((1 + 1/modelVe[4]^2)^2)
                modelVe[6] = Momentum(modelVe[3],0.95)
                modelVe[4] += 1
                modelVe[5] /= modelVe[4]
            end
            @save "$(modelName)_b.bson" modelVe
            if (prevLoss - Base.invokelatest(modelVe[2],intrain,outtrain))/Base.invokelatest(modelVe[2],intrain,outtrain) < modelVe[5]
                modelVe[3] *= (1 + 1/modelVe[4]^2)
                modelVe[6] = Momentum(modelVe[3],0.95)
                println("Increasing Train rate to: $(modelVe[3])")
            end
        end
    else
        for epoch ∈ 0:checklen:(epochs-checklen)
            prevLoss = Base.invokelatest(modelVe[2],intrain,outtrain)
            for epoc ∈ 1:checklen
                train!(modelVe[2], Flux.params(modelVe[1][1:end]), data, modelVe[6])
            end
            if Base.invokelatest(modelVe[2],intrain,outtrain) > prevLoss
                @load "$(modelName)_b.bson" modelVe
                modelVe[3] /= ((1 + 1/modelVe[4]^2)^2)
                modelVe[6] = Momentum(modelVe[3],0.95)
                modelVe[4] += 1
                modelVe[5] /= modelVe[4]
            end
            @save "$(modelName)_b.bson" modelVe
            if (prevLoss - Base.invokelatest(modelVe[2],intrain,outtrain))/Base.invokelatest(modelVe[2],intrain,outtrain) < modelVe[5]
                modelVe[3] *= (1 + 1/modelVe[4]^2)
                modelVe[6] = Momentum(modelVe[3],0.95)
            end
        end
    end
    modelVe[3] /= (1 + 1/modelVe[4]^2)
    #println("test = $(Base.invokelatest(modelVe[2],intrain,outtrain))")
    if save; @save "$(modelName)$(label).bson" modelVe; end
    return modelVe, Base.invokelatest(modelVe[2],intrain,outtrain)
end

function createModelVec(insize, outsize, data; trainrate = 0.0000001, nval = 1, trainfrac = 0.05, save = true, label = "", printstat = false)
    perf = 2000
    model = createModel(insize,outsize)
    modelVec = [model, (x,y) -> sum(Flux.Losses.mse.(model.(x),y)), trainrate, nval, trainfrac, Momentum(trainrate,0.95)]
    while perf > 1000
        model = createModel(insize,outsize)
        modelVec[1] = model
        modelVec[2] = (x,y) -> sum(Flux.Losses.mse.(model.(x),y))
        modelVec[3] = trainrate
        modelVec[4] = nval
        modelVec[5] = trainfrac
        modelVec[6] = Momentum(modelVec[3],0.95)
        modelVec, perf = trainepochs(modelVec, 2000, data;label = label, save = save, printstat = printstat)
    end
    return modelVec
end

function saveModelVec(modelVec,modelName)
    @save "$(modelName).bson" modelVec
    return modelVec
end

function loadModel(modelName)
    @load "$(modelName).bson" modelVec
    return modelVec
end






data, intrain, outtrain = importData("/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/",0,"3273")
insize = size(intrain[1])[1]
outsize = size(outtrain[1])[1]

modelVec = createModelVec(insize,outsize,data)

testData, inTest, outTest = importData("/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/",1,"3273")

println(Base.invokelatest(modelVec[2],inTest,outTest))
for i ∈ 1:5
    global modelVec
    modelVec, perf = trainepochs(modelVec, 4000, data;label = "_T")
    println(Base.invokelatest(modelVec[2],inTest,outTest))
end

saveModelVec(modelVec,"model_v1")

#println((modelVec[1].(inTest) .- outTest)./outTest)
println(sum(((modelVec[1].(inTest)) .- (outTest))./(outTest))/length(outTest))
println(sum(map(x -> map(y -> abs(y),x),((modelVec[1].(inTest)) .- (outTest))./(outTest)))/length(outTest))


println(sum(map(a -> map(b -> abs(b),a),(map(x -> map(y -> exp(y),x),(modelVec[1].(inTest))) .- map(x -> map(y -> exp(y),x),outTest)) ./ (map(x -> map(y -> exp(y)-1,x),outTest))))/length(outTest))