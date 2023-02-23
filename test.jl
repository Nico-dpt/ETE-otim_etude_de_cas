# using CSV, DataFrames

# data = CSV.read("results_final.csv", DataFrame)
# k = 1
date = XLSX.readdata(data_file, "conso_prodfatal", "A"*string(2+k*672)*":A"*string(675+k*672))
heure = XLSX.readdata(data_file, "conso_prodfatal", "B"*string(2+k*672)*":B"*string(675+k*672))

# Tmax = 674

# for t in 1:Tmax
#     data[675+t, 1] = date[t]
#     data[675+t, 2] = heure[t]
# end

# CSV.write("results_final.csv", data)
using CSV, DataFrames

df = CSV.read("results_final.csv", DataFrame ; header =true)
df = df[:, 1:2]
print(df)
push!(df, ["2020-01-01", "00:00:03"], promote=true)
