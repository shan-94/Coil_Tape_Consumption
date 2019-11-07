import math
import xlrd
import xlsxwriter
import os

data_file_name = "\\Coil Data.xlsx"
data_loc = os.getcwd()
data_file = data_loc + data_file_name

data = [] 

coil_set_name = ''
lro = 0
le = 0
dw = 0
ags = 0
no_of_straight_fm_turns = 0
no_of_overhang_fm_turns = 0
no_of_straight_gm_turns = 0
no_of_overhang_gm_turns = 0
total_width = 0.0
total_thickness = 0.0
perimeter = 0.0

# Tape overlap % and tape width - all taken on the maximum side
tape_width = 25
fm_tape_overlap = 0.48
gm_tape_overlap = 0.48
ags_tape_overlap = 0.5
armor_tape_overlap  = 0.5

    ## Formula to calculate consumption ##
def gm_consumption(perimeter, n_straight_layers, n_overhang_layers):
    straight_portion = 2 * n_straight_layers * (perimeter/(tape_width*(1-gm_tape_overlap))) * (lro+140)
    overhang_port_1 = 2 * n_overhang_layers * (perimeter/(tape_width*(1-gm_tape_overlap))) * (le-lro)
    overhang_port_2 = n_overhang_layers * (perimeter/(tape_width*(1-gm_tape_overlap))) * (math.pi*(2*total_thickness+dw))
    overhang_portion = overhang_port_1 + overhang_port_2
    total_gm = straight_portion + overhang_portion 
    return total_gm

def fm_consumption(perimeter, n_straight_layers, n_overhang_layers):
    straight_portion = 2 * n_straight_layers * (perimeter/(tape_width*(1-fm_tape_overlap))) * (lro+140)
    overhang_port_1 = 2 * n_overhang_layers * (perimeter/(tape_width*(1-fm_tape_overlap))) * (le-lro)
    overhang_port_2 = n_overhang_layers * (perimeter/(tape_width*(1-fm_tape_overlap))) * (math.pi*(2*total_thickness+dw))
    overhang_portion = overhang_port_1 + overhang_port_2
    total_fm = straight_portion + overhang_portion 
    return total_fm

def armor_consumption(perimeter):
    armor_1 = 2 * (perimeter/(tape_width*(1-armor_tape_overlap))) * (le-lro)
    armor_2 = (perimeter/(tape_width*(1-armor_tape_overlap))) * (math.pi*(2*total_thickness+dw))
    total_armor = 1.6 * (armor_1 + armor_2)
    return total_armor

def ags_consumption(perimeter):
    total_ags = 2 * (perimeter/(tape_width*(1-ags_tape_overlap))) * (2*ags)
    return total_ags

def get_data_and_calculate_consumption():
    ctr = 0
    wb = xlrd.open_workbook(filename=data_file)
    sheet = wb.sheet_by_index(0)

    no_of_rows = sheet.nrows
    no_of_cols = sheet.ncols

    for i in range(sheet.nrows):
        ctr += 1
        for j in range(sheet.ncols):
            if (i != 0):
                data.append(sheet.cell_value(i,j))
        calc_consumption(data, ctr)
        data.clear()

def calc_consumption(data_list, counter):  
    if (len(data_list) != 0):
        coil_set_name = data_list[0]
        lro = data_list[1]
        le = data_list[2]
        dw = data_list[3]
        ags = data_list[4]
        no_of_straight_fm_turns = data_list[5]
        no_of_overhang_fm_turns = data_list[6]
        no_of_straight_gm_turns = data_list[7]
        no_of_overhang_gm_turns = data_list[8]
        total_width = data_list[8]
        total_thickness = data_list[10]
        perimeter = 2*(total_width + total_thickness)

        print(coil_set_name)
        print(lro)
        print(le)
        print(dw)
        print(ags)
        print(no_of_straight_fm_turns)
        print(no_of_overhang_fm_turns)
        print(no_of_straight_gm_turns)
        print(no_of_overhang_gm_turns)
        print(total_width)
        print(total_thickness)
        print(perimeter)

        total_gm_consumption = gm_consumption(perimeter, no_of_straight_gm_turns, no_of_overhang_gm_turns)/1000
    
        # FM consumption calculation
        # total_fm_consumption = fm_consumption(perimeter, no_of_straight_fm_turns, no_of_overhang_fm_turns)/1000
    
        # #overhang armor calculation
        # total_armor_consumption = armor_consumption(perimeter)/1000

        # #ags calculation
        # total_ags_consumption = ags_consumption(perimeter)/1000

        print("Tape Consumption for ", coil_set_name, "is: ")
        print("GM consumption: ",total_gm_consumption, "m" )
        # print("FM consumption: ",total_fm_consumption, "m" )
        # print("Armor consumption: ",total_armor_consumption, "m" )
        # print("AGS consumption: ",total_ags_consumption, "m\n" )

    return 0    

    return 0

def main():

    #perimeter = total_conduct_perimeter(conductor_width,conductor_thickness,N_turns,parallels)
    # perimeter = 2 * (total_width + total_thickness)

    get_data_and_calculate_consumption()
    # GM consumption calculation
    # total_gm_consumption = gm_consumption(perimeter, no_of_straight_gm_turns, no_of_overhang_gm_turns)/1000
    
    # # FM consumption calculation
    # total_fm_consumption = fm_consumption(perimeter, no_of_straight_fm_turns, no_of_overhang_fm_turns)/1000
    
    # #overhang armor calculation
    # total_armor_consumption = armor_consumption(perimeter)/1000

    # #ags calculation
    # total_ags_consumption = ags_consumption(perimeter)/1000

    # # workbook = xlsxwriter.Workbook('hello.xlsx')
    # # worksheet = workbook.add_worksheet()
    # # worksheet.write('A1', 'Hello World')
    # # workbook.close()
    # print("\n")
    # print("Tape Consumption for ", coil_set_name, "is: ")
    # print("GM consumption: ",total_gm_consumption, "m" )
    # print("FM consumption: ",total_fm_consumption, "m" )
    # print("Armor consumption: ",total_armor_consumption, "m" )
    # print("AGS consumption: ",total_ags_consumption, "m" )

    return 0

if __name__ == '__main__':
    main()





