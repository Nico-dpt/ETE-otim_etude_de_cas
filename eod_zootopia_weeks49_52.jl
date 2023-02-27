#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

Tmax = 673 #optimization for 1 month (4 semaines + 1er pas horaire)
data_file = "Donnees.xlsx"
limit_condition_file = "results_final.csv"
stock_hydro_limit_condition = 0.7

#date et heure
date = XLSX.readdata(data_file, "conso_prodfatal", "A8066:A8738")
heure = XLSX.readdata(data_file, "conso_prodfatal", "B8066:B8738")

#data for load and fatal generation 
load = XLSX.readdata(data_file, "conso_prodfatal", "C8066:C8738")
wind = XLSX.readdata(data_file, "conso_prodfatal", "D8066:D8738")
solar = XLSX.readdata(data_file, "conso_prodfatal", "E8066:E8738")
hydro_fatal = XLSX.readdata(data_file, "conso_prodfatal", "F8066:F8738")
thermal_fatal = XLSX.readdata(data_file, "conso_prodfatal", "G8066:G8738")
#total of RES
P_fatal = wind + solar + hydro_fatal + thermal_fatal

#data for thermal clusters
Nth = 21 #number of thermal generation units
names = XLSX.readdata(data_file, "Thermal_cluster", "B2:B22")
dict_th = Dict(i=> names[i] for i in 1:Nth)
costs_th = XLSX.readdata(data_file, "Thermal_cluster", "I2:I22") # euro/MWh
Pmin_th = XLSX.readdata(data_file, "Thermal_cluster", "G2:G22") # MW
Pmax_th = XLSX.readdata(data_file, "Thermal_cluster", "F2:F22") # MW
dmin = XLSX.readdata(data_file, "Thermal_cluster", "H2:H22") # hours


#data for hydro reservoir
Nhy = 1 #number of hydro generation units
Pmin_hy = zeros(Nhy)
Pmax_hy = XLSX.readdata(data_file, "Parc_electrique", "E20") *ones(Nhy) #MW
cost_hydro = XLSX.readdata(data_file, "Parc_electrique", "H20")*ones(Nhy) # vaut 0 ici 
stock_total_hydro = XLSX.readdata(data_file, "Stock_hydro", "B1")*1000000 #MWh
stock_hydro_initial = XLSX.readdata(data_file, "Stock_hydro", "F3")*ones(Nhy)
apport_hydro = XLSX.readdata(data_file, "historique_hydro", "S8066:S8738") #MWh
stock_max_hy_soft = XLSX.readdata(data_file, "Stock_hydro", "O8068:O8740") #MWh 
stock_min_hy_soft = XLSX.readdata(data_file, "Stock_hydro", "M8068:M8740") #MWh
stock_max_hy_hard = 1.05*stock_max_hy_soft
stock_min_hy_hard = 0.95*stock_min_hy_soft


#costs
cth = repeat(costs_th', Tmax) #cost of thermal generation €/MWh
chy = repeat(cost_hydro, Tmax) #cost of hydro generation €/MWh

cuns = 5000*ones(Tmax) #cost of unsupplied energy €/MWh
cexc = 0*ones(Tmax) #cost of in excess energy €/MWh
#data for STEP/battery
#weekly STEP
Pmax_STEP = XLSX.readdata(data_file, "Parc_electrique", "E21") #MW
rSTEP = XLSX.readdata(data_file, "Parc_electrique", "K21")
stock_volume_STEP = XLSX.readdata(data_file, "Parc_electrique", "L21")


#battery
Pmax_battery = 280 #MW
rbattery = 0.85
d_battery = 2 #hours


#############################
#create the optimization model
#############################
model = Model(HiGHS.Optimizer)

#############################
#define the variables
#############################
#thermal generation variables
@variable(model, Pth[1:Tmax,1:Nth] >= 0)
@variable(model, UCth[1:Tmax,1:Nth], Bin)
@variable(model, UPth[1:Tmax,1:Nth], Bin)
@variable(model, DOth[1:Tmax,1:Nth], Bin)
#hydro generation variables
@variable(model, Phy[1:Tmax,1:Nhy] >= 0)
@variable(model, stock_hydro[1:Tmax,1:Nhy] >=0)
@variable(model, stock_max_depasse_horaire[1:Tmax,1:Nhy] >=0) # Stock que l'on s'autorise à ne pas prélever en plus 
@variable(model, stock_min_depasse_horaire[1:Tmax,1:Nhy] >=0) # Stock que l'on s'autorise à prélever en plus 
@variable(model, num_violations_max[1:Nhy], Int)
@variable(model, num_violations_min[1:Nhy], Int)


#unsupplied energy variables
@variable(model, Puns[1:Tmax] >= 0)
#in excess energy variables
@variable(model, Pexc[1:Tmax] >= 0)
#weekly STEP variables
@variable(model, Pcharge_STEP[1:Tmax] >= 0)
@variable(model, Pdecharge_STEP[1:Tmax] >= 0)
@variable(model, stock_STEP[1:Tmax+1] >= 0)
#battery variables
@variable(model, Pcharge_battery[1:Tmax] >= 0)
@variable(model, Pdecharge_battery[1:Tmax] >= 0)
@variable(model, stock_battery[1:Tmax+1] >= 0)


# #############################
#define the objective function
#############################
@objective(model, Min, sum(Pth.*cth) + sum(Phy.*chy) + Puns'cuns + Pexc'cexc)

#############################
#define the constraints 
#############################
#balance constraint
@constraint(model, balance[t in 1:Tmax], sum(Pth[t,g] for g in 1:Nth) + sum(Phy[t,h] for h in 1:Nhy) + P_fatal[t] + Pdecharge_STEP[t] - Pcharge_STEP[t] +Pdecharge_battery[t] - Pcharge_battery[t] + Puns[t] - load[t] - Pexc[t] == 0)


# INITIAL constraint
# read data condition initial/finale 
data_limit_condition = CSV.read(limit_condition_file, DataFrame , header = true, delim = ";") 


## THERMIQUE
#thermique initial
Pth_initial = data_limit_condition[8065, 3:23]
@constraint(model, initial_Pth[g in 1:Nth], Pth[1,g] == Pth_initial[g])
#thermal unit Pmax constraints
@constraint(model, max_th[t in 1:Tmax, g in 1:Nth], Pth[t,g] <= Pmax_th[g]*UCth[t,g])
#thermal unit Pmin constraints
@constraint(model, min_th[t in 1:Tmax, g in 1:Nth], Pmin_th[g]*UCth[t,g] <= Pth[t,g])
#thermal unit Dmin constraints
for g in 1:Nth
        if (dmin[g] > 1)
            @constraint(model, [t in 2:Tmax], UCth[t,g]-UCth[t-1,g]==UPth[t,g]-DOth[t,g],  base_name = "fct_th_$g")
            @constraint(model, [t in 1:Tmax], UPth[t]+DOth[t]<=1,  base_name = "UPDOth_$g")
            @constraint(model, UPth[1,g]==0,  base_name = "iniUPth_$g")
            @constraint(model, DOth[1,g]==0,  base_name = "iniDOth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] >= sum(UPth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminUPth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] <= 1 - sum(DOth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminDOth_$g")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] >= sum(UPth[i,g] for i in 1:t), base_name = "dminUPth_$(g)_init")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] <= 1-sum(DOth[i,g] for i in 1:t), base_name = "dminDOth_$(g)_init")
    end
end


##########################################################################################
## HYDRO
#hydro unit constraints
@constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t,h] <= Pmax_hy[h])

#hydro stock constraint
@constraint(model, stock_hydro_final[h in 1:Nhy], stock_hydro[Tmax,h] == stock_hydro_limit_condition * stock_total_hydro) #stock final = stock initial
@constraint(model, stock_hydro_actual[h in 1:Nhy,t in 2:Tmax], stock_hydro[t,h] == stock_hydro[t-1,h] - Phy[t-1,h] + apport_hydro[t-1,h]) #contrainte liant stock, turbinage et apport
@constraint(model, hard_stock_max_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] <= stock_max_hy_hard[h])
@constraint(model, hard_stock_min_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] >= stock_min_hy_hard[h])

# hydro initial
Phy_initial = data_limit_condition[8065, 24]
stock_hydro_initial = data_limit_condition[8065, 25]
@constraint(model, initial_Phy[h in 1:Nhy], Phy[1,h] == Phy_initial[h])
@constraint(model, initial_stock_hy[h in 1:Nhy], stock_hydro[1,h] == stock_hydro_initial[h])

#last STEP
@constraint(model, last_step_hydro[h in 1:Nhy], Phy[Tmax,h] == Pmin_hy[h]) 




#########################################################################################
#weekly STEP
@constraint(model,turbinage_max[t in 1:Tmax], Pdecharge_STEP[t]<= Pmax_STEP)
@constraint(model,pompage_max[t in 1:Tmax], Pcharge_STEP[t]<= Pmax_STEP)
for i in 1:4 #modélisation des STEP sur chaque semaine
    @constraint(model, stock_STEP[1+(i-1)*168] == 0, base_name = "stock_initial_semaine_$i") #stock initial = 0 pour chaque semaine
    @constraint(model,[t in (1+(i-1)*168):(168 +(i-1)*168)], stock_STEP[t] <= stock_volume_STEP, base_name = "stock_step_MAX_semaine_$i")
    @constraint(model,[t in (2+(i-1)*168):(168 +(i-1)*168+1)], stock_STEP[t] == stock_STEP[t-1] + Pcharge_STEP[t-1]*rSTEP - Pdecharge_STEP[t-1], base_name = "stock_actuel_STEP_semaine_$i")
end
@constraint(model,stock_STEP_initial_last_step, stock_STEP[Tmax] == stock_STEP[1])
@constraint(model,last_step_STEP_Pturb, Pdecharge_STEP[Tmax] ==0)
@constraint(model,last_step_STEP_Ppomp, Pcharge_STEP[Tmax] ==0)


#contrainte sur le dernier pas de temps
@constraint(model,turbinage_max_final, Pdecharge_STEP[Tmax]<= Pmax_STEP)
@constraint(model,pompage_max_final, Pcharge_STEP[Tmax]<= Pmax_STEP)
@constraint(model,stock_max_STEP_final, stock_STEP[Tmax] <= stock_volume_STEP)


#########################################################################################
#battery
# initial
charge_battery_initial = data_limit_condition[8065 ,29]
decharge_battery_initial = data_limit_condition[8065, 30]
stock_battery_initial = data_limit_condition[8065, 31]
@constraint(model, initial_charge_battery, Pcharge_battery[1] == charge_battery_initial)
@constraint(model, initial_decharge_battery, Pdecharge_battery[1] == decharge_battery_initial)
@constraint(model, initial_stock_battery, stock_battery[1] == stock_battery_initial)

#classic
@constraint(model,battery_decharge_max[t in 1:Tmax], Pdecharge_battery[t]<= Pmax_battery)
@constraint(model,battery_charge_max[t in 1:Tmax], Pcharge_battery[t]<= Pmax_battery)
@constraint(model,stock_max_battery[t in 1:Tmax], stock_battery[t] <= Pmax_battery*d_battery)
@constraint(model,stock_actuel_battery[t in 2:Tmax+1], stock_battery[t] == stock_battery[t-1] + Pcharge_battery[t-1]*rbattery - Pdecharge_battery[t-1]/rbattery)
@constraint(model,stock_initial_final_battery, stock_battery[Tmax] == stock_battery_initial)


#last battery
@constraint(model,last_step_battery_Pturb, Pdecharge_battery[Tmax] ==0)
@constraint(model,last_step_battery_Ppomp, Pcharge_battery[Tmax] ==0)


#no need to print the model when it is too big
#solve the model
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)



#exports results as csv file
th_gen = abs.(value.(Pth))
hy_gen = abs.(value.(Phy))
STEP_charge = abs.(value.(Pcharge_STEP))
STEP_decharge = abs.(value.(Pdecharge_STEP))
STEP_stock = abs.(value.(stock_STEP))
battery_charge = abs.(value.(Pcharge_battery))
battery_decharge = abs.(value.(Pdecharge_battery))
battery_stock = abs.(value.(stock_battery))
hydro_stock = abs.(value.(stock_hydro))



# write result
df = CSV.read("results_final.csv", DataFrame , header =true, delim =";")
print(size(df))
delete!(df,[8065, 8066])

for t in 1:Tmax 
    push!(df, vcat(date[t],
                   heure[t],
                   th_gen[t,:],
                   hy_gen[t,:],
                   hydro_stock[t,:],
                   STEP_charge[t],
                   STEP_decharge[t],
                   STEP_stock[t], 
                   battery_charge[t], 
                   battery_decharge[t], 
                   battery_stock[t], 
                   P_fatal[t], 
                   load[t], 
                   load[t]-P_fatal[t]), promote=true)
end
delete!(df,[8737])
CSV.write("results_final.csv", df, delim=';')
