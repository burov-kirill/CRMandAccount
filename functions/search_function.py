import glob
import itertools
import os
from settings.logs import log


def create_projects_list(values):
    log.info(f'Создание словаря проектов')
    project_list = values['prj_list']
    DDU_files = glob.glob(os.path.join(values['AccPayFld'], "*.xlsx"))
    DKP_files = glob.glob(os.path.join(values['AccSalesFld'], "*.xlsx"))
    CRM_files = glob.glob(os.path.join(values['CRMFld'], "*.xlsx"))
    NEW_ROWS_files = []
    if not values['--CREATE_FILE--']:
        SUMMARY_files = glob.glob(os.path.join(values['SummaryFld'], "*.xlsb"))
        if values['--ADD_STRING--']:
            NEW_ROWS_files = glob.glob(os.path.join(values['new_data_fld'], "*.xlsx"))
    res = []
    keys = ['--FROM_PERIOD--', '--TO_PERIOD--', '--ADD_STRING--', '--CREATE_FILE--', '--REVIEW--', 'save_folder']
    for prj in project_list:
        count_file = 0
        prj_dict = dict()
        if values['--CREATE_FILE--']:
            log.info('\n')
            if prj == 'СПУТНИК':
                prj_dict['spt'] = values['spt']
                log.info(f'Для проекта {prj} данные по номенклатуре находятся по следуюущему пути: \n{values["spt"]}')
            for ddu_file, dkp_file, crm_file in itertools.zip_longest(DDU_files, DKP_files, CRM_files, fillvalue=''):
                if prj in ddu_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} карточка 76 располагается по следующему пути: \n{ddu_file}')
                    prj_dict['AccPay'] = ddu_file
                    count_file+=1
                if prj in dkp_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} карточка 90 располагается по следующему пути: \n{dkp_file}')
                    prj_dict['AccSales'] = dkp_file
                if prj in crm_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} данные CRM располагаются по следующему пути: \n{crm_file}')
                    prj_dict['CRM'] = crm_file
                    count_file += 1
            if count_file >= 2:
                for key in keys:
                    prj_dict[key] = values[key]
                prj_dict['prj'] = prj
                if prj != 'СПУТНИК':
                    prj_dict['spt'] = ''
                res.append(prj_dict)
            else:
                log.info(f'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
                log.info(f'Проект {prj} не включен в сборку')
        else:
            log.info('\n')
            if prj == 'СПУТНИК':
                prj_dict['spt'] = values['spt']
                log.info(f'Для проекта {prj} данные по номенклатуре находятся по следуюущему пути: \n{values["spt"]}')
            for ddu_file, dkp_file, crm_file, sum_file, new_str_file in itertools.zip_longest(DDU_files, DKP_files, CRM_files, SUMMARY_files,NEW_ROWS_files, fillvalue=''):
                if prj in ddu_file.upper().replace('_', ' '):
                    prj_dict['AccPay'] = ddu_file
                    log.info(f'Для проекта {prj} карточка 76 располагается по следующему пути: \n{ddu_file}')
                    count_file+=1
                if prj in dkp_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} карточка 90 располагается по следующему пути: \n{dkp_file}')
                    prj_dict['AccSales'] = dkp_file
                if prj in crm_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} данные CRM располагаются по следующему пути: \n{crm_file}')
                    prj_dict['CRM'] = crm_file
                    count_file += 1
                if prj in sum_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} данные сводный файл располагается по следующему пути: \n{sum_file}')
                    prj_dict['SummaryFile'] = sum_file
                    count_file += 1
                if prj in new_str_file.upper().replace('_', ' '):
                    log.info(f'Для проекта {prj} данные с новыми строками располагются по следующему пути: \n{new_str_file}')
                    prj_dict['new_data'] = new_str_file
                    count_file += 1
            if count_file >= 3 or (count_file>=1 and values['--REVIEW--']):
                for key in keys:
                    prj_dict[key] = values[key]
                prj_dict['prj'] = prj
                if prj != 'СПУТНИК':
                    prj_dict['spt'] = ''
                res.append(prj_dict)
            else:
                log.info(f'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
                log.info(f'Проект {prj} не включен в сборку')
    return res

# print(create_projects_list(PROJECTS, DDU_path, DKP_path, CRM_path))