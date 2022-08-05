import XLSX
import DataFrames
using Flux, Plots
using DataAPI
using BSON: @load
using BSON: @save
using Flux: train!
import Base

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

function trainepochs(modelV, epochs, data; label = "", save = true)
    # How often (in epochs) the model creates a checkpoint and modifies training rate
    checklen = 1000
    modelVe = deepcopy(modelV)
    intrain = data[1][1]
    outtrain = data[1][2]
    @save "$(modelName)_b.bson" modelVe
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
    modelVe[3] /= (1 + 1/modelVe[4]^2)
    #println("test = $(Base.invokelatest(modelVe[2],intrain,outtrain))")
    if save; @save "$(modelName)$(label)_0.bson" modelVe; end
    return modelVe, Base.invokelatest(modelVe[2],intrain,outtrain)
end


# Work Path
#ASMPaths = ["C:\\Users\\anavarro\\Box\\MECS_AMS\\ASM2018-2020.xlsx"]
cd("C:\\Users\\anavarro\\Box\\MECS_AMS")

# Personal Comp Path
#ASMPaths = ["/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/ASM2018-2020.xlsx", "/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/ASM2015-2016.xlsx", "/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/ASM2013-2014.xlsx"]
#cd("/Users/alex/Library/CloudStorage/Box-Box/MECS_AMS/")
ASMPaths = []
for file ∈ readdir()
    if length(file) ≥ 10 && file[1:9] == "TRAIN_ASM"
        push!(ASMPaths, file)
    end
end
println(ASMPaths)

NAICSSector = "3273" # Cement
#NAICSSector = "325" # Chem
#NAICSSector = "3251" # Chemical
#NAICSSector = "32"
#=
includeVec = [   1, 2, 3, 4, 5, 6, 7, 8, 9,10,
                11,12,13,14,15,16,17,18,19,20,
                21,22,23,24,25,26,27,28,29,30,
                31,32,33,34,35,36,37,38,39,40,
                41,42,43,44,45,46,47,48,49,50
]

includeVec = [   2, 3, 9,10,
                11,12,13,16,18,
                22,23,24,25,26,27,28,29,30,
                36,37,38,
                47,48,49,50
]
=#
includeVec =  [2,3,9,10,12,13,16,18,19,20,22,26,30,36,38,47,48]
#=
includeVec = [   2, 3, 4, 5, 7, 8, 9,10,
                11,12,13,
                21,22,23,24,25,26,27,28,29,
                36,
                43,46,47,48,50
]
=#
#includeVec = [2,3,9,10,11,12,13,15,16,18,19,20,21,22,26,30,32,35,50]
#includeVec = [2,3,9,10,12,13,21,22,30]

#includeVec = [i for i ∈ 1:50]


dataVec = []
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
            elseif dataVec == [] && n>1 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICSSector) && "$(r["B"])"[1:length(NAICSSector)] == NAICSSector
                global dataVec = [formatData(r[i]) for i ∈ Columns]
            elseif n>=2 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICSSector) && "$(r["B"])"[1:length(NAICSSector)] == NAICSSector
                dataVec = cat(dataVec,[formatData(r[i]) for i ∈ Columns],dims = 2)
            end
            n += 1
        end
    end
end
dataVec = map(x -> typeof(x) <: Number ? x : 0, dataVec)
dataVec = [i == 1 ? dataVec[i,j]-2000  : log(dataVec[i,j]+1) for i ∈ 1:size(dataVec,1), j ∈ 1:size(dataVec,2)]
dataF = DataFrames.DataFrame(dataVec', labels)
inDataF = DataFrames.select(dataF,includeVec)
outDataF = DataFrames.select(dataF,[51])
insize = size(inDataF,2)
outsize = size(outDataF,2)

intrain = [[inDataF[i, j] for j ∈ 1:insize] for i ∈ 1:size(dataVec,2)]
outtrain = [[outDataF[i, j] for j ∈ 1:outsize] for i ∈ 1:size(dataVec,2)]

data = [(intrain, outtrain)]


global modelName = "model_7"
using BSON: @load

try
    @load "$(modelName)_0.bson" modelVe
    global modelVec = deepcopy(modelVe)
catch er
    println(er)
    local model = createModel(insize,outsize)
    local trainrate = 0.0000001
    local n = 1
    local fract = 0.05
    local predict = y -> model.(y)
    local loss = (x,y) -> sum(Flux.Losses.mse.(predict(x),y))
    local opt = Momentum(trainrate,0.95)
    global modelVec = [model, loss, trainrate, n, fract, opt]
    modelVec, perf = trainepochs(modelVec, 1000, data;label = "_T")
    while perf > 1000
        model = createModel(insize,outsize)
        modelVec[1] = model
        modelVec[2] = (x,y) -> sum(Flux.Losses.mse.(model.(x),y))
        modelVec[3] = 0.0000001
        modelVec[4] = 1
        modelVec[5] = 0.05
        modelVec[6] = Momentum(modelVec[3],0.95)
        modelVec, perf = trainepochs(modelVec, 2000, data;label = "_T")
    end
end

#println(model([inDataF[2,i] for i ∈ 1:8]))
#Optimise.train!(loss, inDataF, , )
#println([model([inDataF[i, j] for j ∈ 1:8])[1] for i ∈ 1:size(dataVec,2)])
#println([inDataF[i, 1] for i ∈ 1:size(dataVec,2)])


global perf
global modelVec
global predict

###### PREPARE TESTING DATA ######

TESTPaths = []
for file ∈ readdir()
    if length(file) ≥ 10 && file[1:8] == "TEST_ASM"
        push!(TESTPaths, file)
    end
end
println(TESTPaths)

dataTestVec = []
for file ∈ TESTPaths
    XLSX.openxlsx(file, enable_cache=false) do f
        sheet = f["Data"]
        n=1
        for r ∈ XLSX.eachrow(sheet)
            if n == 1
                global labels = [(r[i]) for i ∈ Columns]
            elseif dataTestVec == [] && n>1 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICSSector) && "$(r["B"])"[1:length(NAICSSector)] == NAICSSector
                global dataTestVec = [formatData(r[i]) for i ∈ Columns]
            elseif n>=2 && r["C"] != 2012 && length("$(r["B"])") ≥ length(NAICSSector) && "$(r["B"])"[1:length(NAICSSector)] == NAICSSector
                dataTestVec = cat(dataTestVec,[formatData(r[i]) for i ∈ Columns],dims = 2)
            end
            n += 1
        end
    end
end
dataTestVec = map(x -> typeof(x) <: Number ? x : 0, dataTestVec)
dataTestVec = [i == 1 ? dataTestVec[i,j]-2000  : log(dataTestVec[i,j]+1) for i ∈ 1:size(dataTestVec,1), j ∈ 1:size(dataTestVec,2)]
dataTestF = DataFrames.DataFrame(dataTestVec', labels)
inDataTestF = DataFrames.select(dataTestF,includeVec)
outDataTestF = DataFrames.select(dataTestF,[51])

inTest = [[inDataTestF[i, j] for j ∈ 1:insize] for i ∈ 1:size(dataTestVec,2)]
outTest = [[outDataTestF[i, j] for j ∈ 1:outsize] for i ∈ 1:size(dataTestVec,2)]
TestData = [(inTest, outTest)]

println(Base.invokelatest(modelVec[2],inTest,outTest))
for i ∈ 1:5
    global modelVec
    modelVec, perf = trainepochs(modelVec, 2000, data;label = "_T")
    println(Base.invokelatest(modelVec[2],inTest,outTest))
end

#println((modelVec[1].(inTest) .- outTest)./outTest)
println(sum(((modelVec[1].(inTest)) .- (outTest))./(outTest))/length(outTest))
println(sum(map(x -> map(y -> abs(y),x),((modelVec[1].(inTest)) .- (outTest))./(outTest)))/length(outTest))


println(sum(map(a -> map(b -> abs(b),a),(map(x -> map(y -> exp(y),x),(modelVec[1].(inTest))) .- map(x -> map(y -> exp(y),x),outTest)) ./ (map(x -> map(y -> exp(y)-1,x),outTest))))/length(outTest))