from pyexcel_xls import get_data
from pymongo import MongoClient
import re
import pyexcel as pe
from omppackage.omp_connect import *
from omppackage.parse_omp_rtplan import BrachyPlan
import os
import dicom


def get_patient_dob(patient_id):
    """Get the patient date of birth from OMP database"""
    try:
        while True:
            cases = get_patient_cases(patient_id)
            for case in cases:
                plans = get_plans_from_case(patient_id, case)
                for plan in plans:
                    rt_plan_blob = get_rtplan(patient_id, case, plan)
                    output_full_path = r'omppackage\\data\\rtplan.dcm'
                    write_file(rt_plan_blob[0][1], output_full_path)  # save BLOB to dcm
                    ds_input = dicom.read_file(output_full_path)  # read in dcm
                    # print(ds_input)
                    os.remove(output_full_path)
                    patient_dob = ds_input[0x0010, 0x0030].value
                    # print(patient_dob)
                    if patient_dob:
                        return patient_dob
                        break
            return None
            break
    except (IndexError, TypeError) as err:
        print(err)
        pass

def get_omp_data(patient_id):
    """Query OMP database for list of plans for patient ID"""
    list_of_plans = []
    cases = get_patient_cases(patient_id)
    for case in cases:
        plans = get_plans_from_case(patient_id, case)
        for plan in plans:
            rt_plan_blob = get_rtplan(patient_id, case, plan)
            output_full_path = r'omppackage\\data\\rtplan.dcm'
            try:
                write_file(rt_plan_blob[0][1], output_full_path)  # save BLOB to dcm
                ds_input = dicom.read_file(output_full_path)  # read in dcm
                os.remove(output_full_path)
                list_of_plans.append(BrachyPlan(ds_input))
            except(IndexError, AttributeError, TypeError):
                pass
    return list_of_plans


def get_total_treatment_time(patient_id,case_name,plan_name):
    """Calculate total treatment time (per fraction"""
    rt_plan_blob = get_rtplan(patient_id, case_name, plan_name)
    output_full_path = r'omppackage\\data\\rtplan.dcm'
    write_file(rt_plan_blob[0][1], output_full_path)  # save BLOB to dcm
    ds_input = dicom.read_file(output_full_path)                    # read in dcm
    os.remove(output_full_path)
    try:
        my_plan = BrachyPlan(ds_input)
        return my_plan.total_treatment_time
    except AttributeError:
        pass


def handle_assignment(data_dict, data_dict_key, data_in, data_in_key, i1,i2):
    """Insert data into dictionary"""
    try:
        data_dict[data_dict_key] = data_in[data_in_key][i1][i2]
    except (IndexError, KeyError) as e:
        pass


def handle_assignment_simple(data_dict, data_dict_key, data_in):
    try:
        if data_in:
            data_dict[data_dict_key] = data_in
    except (IndexError, KeyError) as e:
        pass


def handle_assignment_date(data_dict, data_dict_key, data_in, data_in_key, i1,i2):
    try:
        data_dict[data_dict_key] = str(data_in[data_in_key][i1][i2])
    except (IndexError, KeyError) as e:
        pass


regex = '^[A-Za-z][0-9]{0,6}[A-Z].xls[A-Za-z]*$'        #regex to find any xls or xlsx files
count = 0
count_failed = 0
for root, dirs, files in os.walk("."):
    for file in files:
        if re.match(regex,file):
            count += 1
            sheet = pe.get_book(file_name=os.path.join(root, file))
            sheet_name = sheet.sheet_names()[0]
            data = get_data(os.path.join(root, file))
            patient_dict = {}
            handle_assignment(patient_dict,'patient_ID', data, sheet_name,2,2)
            handle_assignment_simple(patient_dict,'patient_dob', get_patient_dob(patient_dict['patient_ID']))
            handle_assignment(patient_dict,'patient_name',data,sheet_name,3,2)
            handle_assignment(patient_dict,'consultant', data,sheet_name,2,6)
            handle_assignment(patient_dict,'ebrt_dose_per_fraction',data,sheet_name,5,2)
            handle_assignment(patient_dict,'ebrt_no_fractions', data,sheet_name,6,2)
            handle_assignment(patient_dict,'ebrt_total_dose_tumour', data,sheet_name,7,3)
            handle_assignment(patient_dict,'ebrt_total_dose_oar',data,sheet_name,7,4)
            try:
                omp_data = get_omp_data(patient_dict['patient_ID'])
            except:
                pass
            insertion_list = []
            insertion_number = 0

            for i in range(2, 5):
                insertion_number += 1
                dictionary_temp = {}
                handle_assignment_simple(dictionary_temp,'insertion_number', insertion_number)
                handle_assignment_date(dictionary_temp,'insertion_date', data,sheet_name,9,i)

                if omp_data:
                    try:
                        for treatment in omp_data:
                            if treatment.treatment_date[0:4]+'-'+treatment.treatment_date[4:6]+'-'+treatment.treatment_date[6:]\
                                    == dictionary_temp['insertion_date']:
                                handle_assignment_simple(dictionary_temp, 'total_treatment_time',
                                                         treatment.total_treatment_time)
                    except:
                        pass

                handle_assignment(dictionary_temp,'imaging', data,sheet_name,10,i)
                handle_assignment(dictionary_temp,'prescribed_dose', data,sheet_name,13,i)
                handle_assignment(dictionary_temp,'prescribed_dose_eqd2', data,sheet_name,14,i)
                handle_assignment(dictionary_temp,'point_a_left', data,sheet_name,15,i)
                handle_assignment(dictionary_temp,'point_a_left_eqd2', data,sheet_name,16,i)
                handle_assignment(dictionary_temp,'point_a_right', data,sheet_name,17,i)
                handle_assignment(dictionary_temp,'point_a_right_eqd2', data,sheet_name,18,i)
                handle_assignment(dictionary_temp,'mean_point_a', data,sheet_name,19,i)
                handle_assignment(dictionary_temp,'mean_point_a_eqd2', data,sheet_name,20,i)
                handle_assignment(dictionary_temp,'hr_ctv_volume_cm3', data,sheet_name,21,i)
                handle_assignment(dictionary_temp,'hr_ctv_d90_gy', data,sheet_name,22,i)
                handle_assignment(dictionary_temp,'hr_ctv_d90_gy_eqd2', data,sheet_name,23,i)
                handle_assignment(dictionary_temp,'hr_ctv_v100_percent', data,sheet_name,24,i)
                handle_assignment(dictionary_temp,'bladder_volume_cm3', data,sheet_name,25,i)
                handle_assignment(dictionary_temp,'bladder_icru_gy', data,sheet_name,26,i)
                handle_assignment(dictionary_temp,'bladder_icru_gy_eqd2', data,sheet_name,27,i)
                handle_assignment(dictionary_temp,'bladder_d2cc_gy', data,sheet_name,28,i)
                handle_assignment(dictionary_temp,'bladder_d2cc_gy_eqd2', data,sheet_name,29,i)
                handle_assignment(dictionary_temp,'rectum_volume_cm3', data,sheet_name,30,i)
                handle_assignment(dictionary_temp,'rectum_icru_gy', data,sheet_name,31,i)
                handle_assignment(dictionary_temp,'rectum_icru_gy_eqd2', data,sheet_name,32,i)
                handle_assignment(dictionary_temp,'rectum_d2cc_gy', data,sheet_name,33,i)
                handle_assignment(dictionary_temp,'rectum_d2cc_gy_eqd2', data,sheet_name,34,i)
                handle_assignment(dictionary_temp,'bowel_volume_cm3', data,sheet_name,35,i)
                handle_assignment(dictionary_temp,'bowel_d2cc_gy', data,sheet_name,36,i)
                handle_assignment(dictionary_temp,'bowel_d2cc_gy_eqd2', data,sheet_name,37,i)

                insertion_list.append(dictionary_temp)

            patient_dict['insertions'] = insertion_list

            client = MongoClient()
            db = client.patient_database
            collection = db.patients

            patients = db.patients
            patients.insert_one(patient_dict)
