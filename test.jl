using CSV, DataFrames
data_limit_condition = "results_final.csv"
data = CSV.read(data_limit_condition, DataFrame ; header = true)
a = data[3,3]
print(a[1])
