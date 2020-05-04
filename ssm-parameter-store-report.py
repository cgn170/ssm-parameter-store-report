import boto3
from collections import defaultdict
import xlsxwriter


def get_resources_from_describe_parameter(ssm_details):
    results = ssm_details['Parameters']
    resources = [result for result in results]
    next_token = ssm_details.get('NextToken', None)
    return resources, next_token


result = defaultdict(lambda: defaultdict(lambda: defaultdict(str)))

#############
# AWS ACCOUNT
##############
try:
    print ("Querying AWS ACCOUNT")
    ssm = boto3.client('ssm',
                       # aws_access_key_id='',
                       # aws_secret_access_key='',
                       region_name="us-east-1")
    next_token = ' '
    ssm_variables = []
    while next_token is not None:
        ssm_details = ssm.describe_parameters(MaxResults=50, NextToken=next_token)
        current_batch, next_token = get_resources_from_describe_parameter(ssm_details)
        ssm_variables += current_batch
except Exception as e:
    print ("Error: {}".format(e))
# Loop for each ssm parameter
for ssm_variable in ssm_variables:
    try:
        ssm_variable_name = ssm_variable.get("Name")
        path_info = ssm_variable_name.split("/")
        env = path_info[2]
        ms_name = path_info[3]
        ms_variable = path_info[4]
        ssm_value = ssm.get_parameter(Name=ssm_variable_name, WithDecryption=True)["Parameter"].get("Value")
        result[ms_name][ms_variable][env] = ssm_value
    except Exception as e:
        print ("Error {0} - querying {1} ".format(e, ssm_variable.get("Name")))


# Writing report
workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet("Variables")

header_format = workbook.add_format({'bg_color': '#93CCEA'})

# Writing headers
worksheet.set_column('A:G', 25)
worksheet.write('A1', 'MICROSERVICE', header_format)
worksheet.write('B1', 'VARIABLE', header_format)
worksheet.write('C1', 'DEV', header_format)
worksheet.write('D1', 'QA', header_format)
worksheet.write('E1', 'STG', header_format)
worksheet.write('F1', 'PRO', header_format)
worksheet.freeze_panes(1, 0)

row = 1
ms_keys_order = result.keys()
ms_keys_order.sort()
for ms_key in ms_keys_order:
    var_keys_order = result[ms_key].keys()
    var_keys_order.sort()
    for var_key in var_keys_order:
        print ("MS: {0}, VARIABLE: {1}, DEV_VALUE: {2}, QA_VALUE: {3}, STG_VALUE: {4}, PROA_VALUE: {5}"
               .format(ms_key, var_key,
                       result[ms_key][var_key].get("dev", "No available"),
                       result[ms_key][var_key].get('qa', "No available"),
                       result[ms_key][var_key].get('stg', " No available"),
                       result[ms_key][var_key].get('pro', " No available"),
                       ))
        worksheet.write(row, 0, ms_key)
        worksheet.write(row, 1, var_key)
        worksheet.write(row, 2, result[ms_key][var_key].get("dev", "No available"))
        worksheet.write(row, 3, result[ms_key][var_key].get("qa", "No available"))
        worksheet.write(row, 4, result[ms_key][var_key].get("stg", "No available"))
        worksheet.write(row, 5, result[ms_key][var_key].get("pro", "No available"))
        row += 1
# Add autofiler
worksheet.autofilter('A1:G{0}'.format(row))

workbook.close()

print('done')
