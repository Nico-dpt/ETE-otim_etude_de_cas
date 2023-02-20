#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

Tmax = 674 #optimization for 1 month (4 semaines + 1er pas horaire)
data_file = "Donnees.xlsx"

#date et heure
date = XLSX.readdata(data_file, "conso_prodfatal", "A2:A675")
heure = XLSX.readdata(data_file, "conso_prodfatal", "B2:B675")

#data for load and fatal generation 
load = XLSX.readdata(data_file, "conso_prodfatal", "C2:C675")
wind = XLSX.readdata(data_file, "conso_prodfatal", "D2:D675")
solar = XLSX.readdata(data_file, "conso_prodfatal", "E2:E675")
hydro_fatal = XLSX.readdata(data_file, "conso_prodfatal", "F2:F675")
thermal_fatal = XLSX.readdata(data_file, "conso_prodfatal", "G2:G675")
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
stock_hydro_initial = XLSX.readdata(data_file, "Stock_hydro", "F3")*ones(Nhy)
apport_hydro = XLSX.readdata(data_file, "historique_hydro", "S2:S675") #MWh
stock_max_hy_soft = XLSX.readdata(data_file, "Stock_hydro", "E4:E677") #MWh 
stock_min_hy_soft = XLSX.readdata(data_file, "Stock_hydro", "C4:C677") #MWh
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

#hydro unit constraints
@constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t,h] <= Pmax_hy[h])
@constraint(model, last_step_hydro[h in 1:Nhy], Phy[Tmax,h] == Pmin_hy[h]) 

#hydro stock constraint
@constraint(model, stoch_hy_initial[h in 1:Nhy], stock_hydro[1,h] == stock_hydro_initial[h]) # stock initial
@constraint(model, stock_hydro_final[h in 1:Nhy], stock_hydro[Tmax,h] == stock_hydro_initial[h]) #stock final = stock initial
@constraint(model, stock_hydro_actual[h in 1:Nhy,t in 2:Tmax], stock_hydro[t,h] == stock_hydro[t-1,h] - Phy[t-1,h] + apport_hydro[t-1,h]) #contrainte liant stock, turbinage et apport
@constraint(model, hard_stock_max_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] <= stock_max_hy_hard[h])
@constraint(model, hard_stock_min_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] >= stock_min_hy_hard[h])


# print(stock_max_depasse_horaire[1,1] > 0)
# @constraint(model, soft_stock_max_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] <= stock_max_hy_soft[h] + stock_max_depasse_horaire[t,h])
# @constraint(model, soft_stock_min_hy[h in 1:Nhy,t in 1:Tmax], stock_hydro[t,h] >= stock_min_hy_soft[h] - stock_min_depasse_horaire[t,h])
# @constraint(model, max_violation[h in 1:Nhy], num_violations_max[h] >= sum([stock_max_depasse_horaire[t,h] > 0 for t in 1:Tmax]))
# @constraint(model, min_violation[h in 1:Nhy],  num_violations_min[h] >= sum([stock_min_depasse_horaire[t,h] > 0 for t in 1:Tmax]))
# @constraint(model, num_violations_max[h in 1:Nhy] <= 20)
# @constraint(model, num_violations_min[h in 1:Nhy] <= 20)


#@constraint(model, nombre_de_depassement_max_hy[h in 1:Nhy], sum(depassement_max_hy[t,h] for t in 1:Tmax) <= 20)
#@constraint(model, nombre_de_depassement_min_hy[h in 1:Nhy], sum(depassement_min_hy[t,h] for t in 1:Tmax) <= 20)
#@constraint(model, depassement_max[h in 1:Nhy,t in 1:Tmax], (1-2*depassement_max_hy[t,h])*(stock_max_hy_soft - stock_hydro[t,h]) >= 0)
#@constraint(model, depassement_min[h in 1:Nhy,t in 1:Tmax], (1-2*depassement_min_hy[t,h])*(stock_min_hy_soft - stock_hydro[t,h]) <= 0)





#weekly STEP
@constraint(model,turbinage_max[t in 1:Tmax], Pdecharge_STEP[t]<= Pmax_STEP)
@constraint(model,pompage_max[t in 1:Tmax], Pcharge_STEP[t]<= Pmax_STEP)
#@constraint(model,stock_initial,stock_STEP[1] ==0)
for i in 1:4 #modélisation des STEP sur chaque semaine
    @constraint(model, stock_STEP[1+(i-1)*168] == 0, base_name = "stock_initial_semaine_$i") #stock initial = 0 pour chaque semaine
    @constraint(model,[t in (1+(i-1)*168):(168 +(i-1)*168)], stock_STEP[t] <= stock_volume_STEP, base_name = "stock_step_MAX_semaine_$i")
    @constraint(model,[t in (2+(i-1)*168):(168 +(i-1)*168+1)], stock_STEP[t] == stock_STEP[t-1] + Pcharge_STEP[t-1]*rSTEP - Pdecharge_STEP[t-1], base_name = "stock_actuel_STEP_semaine_$i")
    #@constraint(model, stock_STEP[1+(i-1)*168] == stock_STEP[1 + i*168], base_name = "stock_initial_final_semaine_$i") #stock initial = stock final #inutile ??    
end
@constraint(model,stock_STEP_initial_last_step, stock_STEP[Tmax] == stock_STEP[1])
@constraint(model,last_step_STEP_Pturb, Pdecharge_STEP[Tmax] ==0)
@constraint(model,last2_step_STEP_Pturb, Pdecharge_STEP[Tmax-1] ==0)
@constraint(model,last_step_STEP_Ppomp, Pcharge_STEP[Tmax] ==0)
@constraint(model,last2_step_STEP_Ppomp, Pcharge_STEP[Tmax-1] ==0)



#contrainte sur le dernier pas de temps
@constraint(model,turbinage_max_final, Pdecharge_STEP[Tmax]<= Pmax_STEP)
@constraint(model,pompage_max_final, Pcharge_STEP[Tmax]<= Pmax_STEP)
@constraint(model,stock_max_STEP_final, stock_STEP[Tmax] <= stock_volume_STEP)

#battery
@constraint(model,battery_decharge_max[t in 1:Tmax], Pdecharge_battery[t]<= Pmax_battery)
@constraint(model,battery_charge_max[t in 1:Tmax], Pcharge_battery[t]<= Pmax_battery)
@constraint(model,stock_max_battery[t in 1:Tmax], stock_battery[t] <= Pmax_battery*d_battery)
@constraint(model,stock_actuel_battery[t in 2:Tmax+1], stock_battery[t] == stock_battery[t-1] + Pcharge_battery[t-1]*rbattery - Pdecharge_battery[t-1]/rbattery)
@constraint(model,stock_initial_final_battery, stock_battery[1] == stock_battery[Tmax])
@constraint(model,stock_initial_battery, stock_battery[1] == 0)

#no need to print the model when it is too big
#solve the model
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)



#exports results as csv file
th_gen = value.(Pth)
hy_gen = value.(Phy)
STEP_charge = value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)
STEP_stock = value.(stock_STEP)
battery_charge = value.(Pcharge_battery)
battery_decharge = value.(Pdecharge_battery)
battery_stock = value.(stock_battery)
hydro_stock = value.(stock_hydro)



# new file created
touch("results_4firstweek.csv")

# file handling in write mode
f = open("results_4firstweek.csv", "w")


write(f,"Date; heure;")

for name in names
    write(f, "$name ;")
end
write(f, "P_hydro; Hydro_stock; STEP pompage ; STEP turbinage; STEP_stock ; Batterie injection ; Batterie soutirage ; Batterie Stock ; P_fatal ; load ; Net load \n")

for t in 1:Tmax
    write(f, "$(date[t]) ; $(heure[t]);")
    for g in 1:Nth
        write(f, "$(th_gen[t,g]) ; ")
    end
    for h in 1:Nhy
        write(f, "$(hy_gen[t,h]); $(hydro_stock[t,h]) ;")
    end
    write(f, "$(STEP_charge[t]) ; $(STEP_decharge[t]); $(STEP_stock[t]) ;")
    write(f, "$(battery_charge[t]) ; $(battery_decharge[t]) ; $(battery_stock[t]) ;")
    write(f, "$(P_fatal[t]) ;  $(load[t]) ; $(load[t]-P_fatal[t]) \n")

end

close(f)
