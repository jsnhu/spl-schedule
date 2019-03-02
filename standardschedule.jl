using JuMP, GLPKMathProgInterface, DataFrames, Taro

Taro.init()

# get staff, shift, preference tables
staff_df = DataFrame(Taro.readxl("table.xlsx", "Dictionary", "A2:C10"))
shift_df = DataFrame(Taro.readxl("table.xlsx", "Dictionary", "D2:G19"))
pref_df  = DataFrame(Taro.readxl("table.xlsx", "PrefMatrix", "A1:Q8", header = false))
#= Note:
Taro.readxl should return a DataFrame,
but it returns an array of
namedtuples instead for some reason
=#

# categorize staff by their type
both_staff = []
desk_staff = []
shel_staff = []

for k in 1:size(staff_df,1)
    if staff_df[k,:Type] == "both"
        push!(both_staff,k)
    elseif staff_df[k,:Type] == "desk"
        push!(desk_staff,k)
    else
        push!(shel_staff,k)
    end
end

# categorize shifts by their day of the week


# categorize shifts by their type



for x in 1:size(staff_df, 1)
    if staff_df[x,3] === "both"
        push!(both_staff,x)
    elseif staff_df[x,3] === "desk"
        push!(desk_staff,x)
    else
        push!(shel_staff,x)
    end
end








# convert to integer matrix
pref_matrix = Array{Int64}(Matrix(pref_df))

# get number of staff and shifts
staff = size(pref_matrix)[1]
shifts = size(pref_matrix)[2]

# optimization model
m = Model(solver = GLPKSolverMIP())

# 8x17 binary assignment matrix
# 1 if employee i assigned to shift j, 0 otherwise
@variable(m, x[1:staff, 1:shifts], Bin)

# maximize preference score sum
@objective(m, Max, sum(pref_matrix[i,j]*x[i,j] for i in 1:staff, j in 1:shifts))

# constraints

# cons1: exactly one person per shift
for j in 1:shifts
    @constraint(m, sum( x[i,j] for i in 1:staff) == 1)
end

#cons2: maximum 4 shifts per week per person
for i in 1:staff
    @constraint(m, sum( x[i,j] for j in 1:shifts) <= 4)
end

#cons3: no employee works both Saturday and Sunday in one weekend
#       shifts 4, 5, 15, 16, 17
for i in 1:staff
    @constraint(m, sum( x[i,j] for j in [4,5,15,16,17]) <= 1)
end

#cons4: employees 3-5 (desk only) cannot work shelving shifts
#       shifts 6-17
for i in 3:5
    @constraint(m, sum( x[i,j] for j in 6:staff) == 0)
end

#cons5: employees 6-8 (shelving only) cannot work desk shifts
#       shifts 1-5
for i in 6:8
    @constraint(m, sum( x[i,j] for j in 1:5) == 0)
end

#cons6: nobody works two shifts in one day:
for i in 1:staff
    @constraint(m,  sum( x[i,j] for j = [1,6,7]) <= 1)      # Mon
    @constraint(m,  sum( x[i,j] for j = [8,9]) <= 1)        # Tue
    @constraint(m,  sum( x[i,j] for j = [10,11]) <= 1)      # Wed
    @constraint(m,  sum( x[i,j] for j = [2,12,13]) <= 1)    # Thu
    @constraint(m,  sum( x[i,j] for j = [3,14]) <= 1)       # Fri
    @constraint(m,  sum( x[i,j] for j = [4,15]) <= 1)       # Sat
    @constraint(m,  sum( x[i,j] for j = [5,16,17]) <= 1)    # Sun
end

# print(m)

status = solve(m)

println("Objective value: ", getobjectivevalue(m))
assn_matrix = Array{Int64}(getvalue(x))

# create dataframe and add rows
assn_df = DataFrame(Employee = Int[], Shift = Int[], Score = Int[])

for i in 1:staff
    for j in 1:shifts
        if assn_matrix[i,j] == 1
            push!(assn_df, (i, j, pref_matrix[i,j]))
        end
    end
end

# print assignments
show(assn_df, allrows = true)
