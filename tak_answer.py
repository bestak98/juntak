from copy import deepcopy as cp
import csv

PRODUCTION_REQUIREMENT = {'A': 20, 'B': 10, 'C': 1}  # 생산 요구량
# SCHEDULE = {
#     'M-1': ['LOT-4_C-1', 'LOT-3_B-1'],
#     'M-2': ['LOT-2_A-3', 'LOT-1_A-3'],
#     'M-3': ['LOT-1_A-1', 'LOT-1_A-2', 'LOT-2_A-1', 'LOT-2_A-2', 'LOT-4_C-2'],
#     'M-4': ['LOT-3_B-2', 'LOT-3_B-3']}

SCHEDULE = { 'M-1': ['A-1', 'B-2'],
'M-2': ['A-3','B-1'],
'M-3': ['B-3', 'A-2'],
'M-4': ['C-2', 'C-1'] }

# 동종 셋업 (job type은 같은데 operation이 다른 경우) = 3/ 이종 셋업 (job type 다른 경우) = 6
last_operation = {'M-1':'A-3','M-2':'A-3','M-3':'B-3','M-4':'C-2'}
setup_time_dict = {'M-1': [], 'M-2': [], 'M-3': [], 'M-4': []}

# 셋업타임 계산
for key, value in SCHEDULE.items():
    machine_num = key
    for job_operation in value:
        if job_operation == value[0]:
            if job_operation[0] == last_operation[machine_num][0]:
                setup_time = 3
            else:
                setup_time = 6
        else:
            if job_operation[0] == value[value.index(job_operation)-1][0]:
                    setup_time = 3
            else:
                    setup_time = 6
        setup_time_dict[key].append(setup_time)
print("setup_time_dict:",setup_time_dict)

# job[A,B,C]구분
job_list = list(PRODUCTION_REQUIREMENT.keys())

# SCHEDULE-> M-1_A-1 꼴로 변형 
new_schedule = cp(SCHEDULE)
for key, value in new_schedule.items():
    for i in range(len(value)):
        try:
            new_schedule[key][i] = key + '_' + new_schedule[key][i]
        except:
            continue

file = open('Job-shop Scheduling.csv', 'r')
file_read = csv.DictReader(file)

# 기계-작업 시간 불러오기
DICT1 = {}
for i in file_read:
    name = i['']
    for j, m in i.items():
        rows = list(i.keys())
        columns = list(i.values())
        del columns[0], rows[0]
        for a in range(len(rows)):
            rows[a] = rows[a] + "_" + name
        dictionary = {number: int(columns[i]) for i, number in enumerate(rows)}
        DICT1.update(dictionary)
        break
print(DICT1)
Remain_schedule = cp(new_schedule)

def FJSP_makespan(SCHEDULE, DICT1):

    # 작업의 개수와 기계의 개수 파악
    num_of_job = (len(PRODUCTION_REQUIREMENT.keys()))
    num_of_machine = len(SCHEDULE.keys())

    # 초기화
    job_completion_time_list = []
    job_last_finshed_operation_index_list = []
    for reset_job in range(num_of_job):
        job_completion_time_list.append(0)
        job_last_finshed_operation_index_list.append(0)

    machine_completion_time_list = []
    for reset_machine in range(num_of_machine):
        machine_completion_time_list.append(0)

    # 계산
    while True:
        is_break = True

        for key, value in Remain_schedule.items():
            try:
                job_index_1 = Remain_schedule[key][0].split("_")[-1]  # job ABC
                job_index = job_index_1.split("-")[0]
                operation_index = int(Remain_schedule[key][0].split("-")[-1]) # op 123
            except:
                continue

            for i in range(len(job_list)):
                if job_index == job_list[i]:
                    job_index = i + 1
                    
            if job_last_finshed_operation_index_list[job_index - 1] == operation_index - 1:
                is_break = False
                machine_index = int(key.split("-")[1])
                processing_time = DICT1[Remain_schedule[key][0]]
                start_time = max(job_completion_time_list[job_index - 1], machine_completion_time_list[machine_index - 1])
                setup_time = setup_time_dict[key][0]
                job_completion_time_list[job_index - 1] = start_time + setup_time + processing_time
                machine_completion_time_list[machine_index - 1] = start_time + setup_time + processing_time
                job_last_finshed_operation_index_list[job_index - 1] += 1  # op가 끝난 것
                print(">>> MACHINE : {} JOB-OPERATION : {}  PROCESSING_TIME : {} SETUP_TIME : {} ".format(key, Remain_schedule[key][0], processing_time, setup_time))
                print("Update", job_last_finshed_operation_index_list, machine_completion_time_list, job_completion_time_list)
                del Remain_schedule[key][0],  setup_time_dict[key][0]
                print(Remain_schedule)

        if is_break:
            break
    # 오류
    for i in range(num_of_machine):
        if len(list(Remain_schedule.values())[i]) != 0:
            print("Error: Check SCHEDULE")
            exit()

    return max((machine_completion_time_list))

print("RESULT:", FJSP_makespan(SCHEDULE, DICT1))