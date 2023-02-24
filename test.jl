# using CSV, DataFrames
using XLSX
# data = CSV.read("results_final.csv", DataFrame)
k = 0
date = XLSX.readdata("Donnees.xlsx", "conso_prodfatal", "A"*string(2+k*672)*":A"*string(675+k*672))
heure = XLSX.readdata("Donnees.xlsx", "conso_prodfatal", "B"*string(2+k*672)*":B"*string(675+k*672))

# Tmax = 674

# for t in 1:Tmax
#     data[675+t, 1] = date[t]
#     data[675+t, 2] = heure[t]
# end

# CSV.write("results_final.csv", data)
using CSV, DataFrames

df = CSV.read("results_final.csv", DataFrame ; header =true)
df = df[:, 1:2]
#delete!(df,[674])
#push!(df, [(date[1]), "00:00:03"], promote=true)
#a = vcat(1; th_gen[1,:])
print(size(df))
