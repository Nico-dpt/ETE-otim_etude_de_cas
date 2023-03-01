a = [1,2,3]
b = [4,5,6]

combinaison = []
for i in 1:length(a)
    for j in 1:length(b)
        if i!=j
            push!(combinaison, [a[i], b[j]])
        end
    end
