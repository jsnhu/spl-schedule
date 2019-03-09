using JuMP, GLPKMathProgInterface, DataFrames, Taro

Taro.init()

# obtain keys (bottom right corner of cell range to read)
staff_cell  = getCell(getRow(getSheet(Workbook("table.xlsx"), "Keys"),1),1)
shift_cell  = getCell(getRow(getSheet(Workbook("table.xlsx"), "Keys"),2),1)
pref_cell   = getCell(getRow(getSheet(Workbook("table.xlsx"), "Keys"),3),1)

staff_key = string("A2:", getCellValue(staff_cell))
shift_key = string("D2:", getCellValue(shift_cell))
pref_key  = string("A1:", getCellValue(pref_cell))


# get staff, shift, preference tables
staff_df = DataFrame(Taro.readxl("table.xlsx", "Dictionary", staff_key))
shift_df = DataFrame(Taro.readxl("table.xlsx", "Dictionary", shift_key))
pref_df  = DataFrame(Taro.readxl("table.xlsx", "PrefMatrix", pref_key, header = false))
#= Note:
Taro.readxl should return a DataFrame,
but it returns an array of
namedtuples instead for some reason
=#

# categorize staff by their type
both_staff = Int64[]
desk_staff = Int64[]
shel_staff = Int64[]

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
mon_shift = Int64[]
tue_shift = Int64[]
wed_shift = Int64[]
thu_shift = Int64[]
fri_shift = Int64[]
sat_shift = Int64[]
sun_shift = Int64[]

for k in 1:size(shift_df,1)
    if shift_df[k,:Day] == "Mon"
        push!(mon_shift,k)
    elseif shift_df[k,:Day] == "Tue"
        push!(tue_shift,k)
    elseif shift_df[k,:Day] == "Wed"
        push!(wed_shift,k)
    elseif shift_df[k,:Day] == "Thu"
        push!(thu_shift,k)
    elseif shift_df[k,:Day] == "Fri"
        push!(fri_shift,k)
    elseif shift_df[k,:Day] == "Sat"
        push!(sat_shift,k)
    else
        push!(sun_shift,k)
    end
end

# categorize shifts by their type

desk_shift = Int64[]
shel_shift = Int64[]

for k in 1:size(shift_df,1)
    if shift_df[k,:Type] == "desk"
        push!(desk_shift,k)
    else
        push!(shel_shift,k)
    end
end

# convert pref table to integer matrix
pref_matrix = Array{Int64}(Matrix(pref_df))

# get number of staff and shifts
staff = size(pref_matrix)[1]
shift = size(pref_matrix)[2]

# optimization model
m = Model(solver = GLPKSolverMIP())

# staff x shift binary assignment matrix
# 1 if employee i assigned to shift j, 0 otherwise
@variable(m, x[1:staff, 1:shift], Bin)

# maximize preference score sum
@objective(m, Max, sum(pref_matrix[i,j]*x[i,j] for i in 1:staff, j in 1:shift))

# constraints

# cons1: exactly one person per shift
for j in 1:shift
    @constraint(m, sum( x[i,j] for i in 1:staff) == 1)
end

# cons2: maximum 4 shifts per week per person
for i in 1:staff
    @constraint(m, sum( x[i,j] for j in 1:shift) <= 4)
end

# cons3: minimum 1 shift per week per person
for i in 1:staff
    @constraint(m, sum( x[i,j] for j in 1:shift) >= 1)
end

# cons4: no employee works both Sat and Sun in one weekend
for i in 1:staff
    @constraint(m, sum( x[i,j] for j in union(sat_shift,sun_shift)) <= 1)
end

# cons5: desk employees cannot work shelving shifts
for i in desk_staff
    @constraint(m, sum( x[i,j] for j in shel_shift) == 0)
end

# cons6: shelving employees cannot work desk shifts
for i in shel_staff
    @constraint(m, sum( x[i,j] for j in desk_shift) == 0)
end

# cons7: nobody works 2 shifts (or more) in one day:
for i in 1:staff
    @constraint(m,  sum( x[i,j] for j in mon_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in tue_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in wed_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in thu_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in fri_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in sat_shift) <= 1)
    @constraint(m,  sum( x[i,j] for j in sun_shift) <= 1)
end

# print(m)

status = solve(m)

println("Objective value: ", getobjectivevalue(m))
assn_matrix = Array{Int64}(getvalue(x))

# create assignment dataframe
assn_df = DataFrame(Employee = Int[], Shift = Int[], Score = Int[])

# read all assigned shifts into dataframe
for i in 1:staff
    for j in 1:shift
        if assn_matrix[i,j] == 1
            push!(assn_df, (i, j, pref_matrix[i,j]))
        end
    end
end

# print assignments
show(assn_df, allrows = true)
