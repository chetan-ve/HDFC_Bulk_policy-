
import time
import pandas as pd
import json
import requests
import math
import openpyxl
from dateutil.relativedelta import relativedelta
from datetime import datetime


full_python_output = []
t1 = time.time()
#-------
# excel_file = 'sample.xlsx'
# sheet_name = 'Sample Policies'

excel_file = r'C:\Users\chetan.p\Downloads\Input_PPR_Generation_Policy_29092026.xlsx'
sheet_name = 'Sheet1'
data_values = pd.read_excel(excel_file, sheet_name=sheet_name) 
policy_lower_range = 0
policy_upper_range = 500
list_bajaj_not_run_cases = []

policy_range = data_values.iloc[policy_lower_range:policy_upper_range]


# api url
url = "https://api-dev.valuenable.in/ul-calc"
for index, row in policy_range.iterrows():
    
    if pd.notna(row["UIN"]) and pd.notna(row["RCD"]) and pd.notna(row["PT"]) and pd.notna(row["PPT"]):      
        uin = str(row["UIN"])
        plan_option = "NA" if pd.isna(row.get("plan_option")) else str(row.get("plan_option")).strip()
        pw= row["pw"]
        rcd = row["RCD"].strftime("%d-%m-%Y")
        pt = int(row["PT"])
        ppt = int(row["PPT"])
        modal_premium = row["Modal Premium"]
        def get_fund_name(n):
            val = row.get("Fund {}".format(n))
            return "" if pd.isna(val) else str(val)

        def get_fund_alloc(n):
            val = row.get("Fund alloc {}".format(n))
            return 0 if pd.isna(val) else (val * 100)

        fund_1 = get_fund_name(1)
        fund_1_alloc = get_fund_alloc(1)
        fund_2 = get_fund_name(2)
        fund_2_alloc = get_fund_alloc(2)
        fund_3 = get_fund_name(3)
        fund_3_alloc = get_fund_alloc(3)
        fund_4 = get_fund_name(4)
        fund_4_alloc = get_fund_alloc(4)
        fund_5 = get_fund_name(5)
        fund_5_alloc = get_fund_alloc(5)
        fund_6 = get_fund_name(6)
        fund_6_alloc = get_fund_alloc(6)
        fund_7 = get_fund_name(7)
        fund_7_alloc = get_fund_alloc(7)
        premium_frequency = int(row["Premium Frequency"])
        gender = row["Gender"]
        fund_value = row["Fund Value"]
        ptd = row["PTD"].strftime("%d-%m-%Y")
        dob = row["DOB"].strftime("%d-%m-%Y")
        sum_assured = row["Sum Assured"]
        valuation_date = row["Valuation date"].strftime("%d-%m-%Y")
        policy_number = str(row["Policy Number"])
        module = "bajaj"
        sr = str(row['Sr no'])
        level_cover_period = int(row["level_cover_period"])
        policy_holder_name = str(row.get("policy_holder_name"))
        # Define the GraphQL query
        query = '''
            query Projection(
                $uin: String!,
                $plan_option: String,
                $pw:Float,
                $rcd: GraphQLDateTime!,
                $ppt: Int!,
                $pt: Int!,
                $modal_premium: Float!,
                $fund_1: String,
                $fund_1_alloc: Float,
                $fund_2: String,
                $fund_2_alloc: Float,
                $fund_3: String,
                $fund_3_alloc: Float,
                $fund_4: String,
                $fund_4_alloc: Float,
                $fund_5: String,
                $fund_5_alloc: Float,
                $fund_6: String,
                $fund_6_alloc: Float,
                $fund_7: String,
                $fund_7_alloc: Float,
                $premium_frequency: Int!,
                $gender: Gender!,
                $fund_value: Float!,
                $ptd: GraphQLDateTime!,
                $dob: GraphQLDateTime!,
                $sum_assured: Float!,
                $valuation_date: GraphQLDateTime!
                $policy_number: String,
                $module: String,
                $level_cover_period:Int,
                $max_projection_year:Int,
                $is_pdf:Boolean,
                $calc_pay_zero:Boolean,
                $calc_pay_one:Boolean,
                $calc_pay_three:Boolean,
                $policy_holder_name:String
            ) {
            Projection(
            uin: $uin,
                    plan_option:$plan_option,
                    pw : $pw
                    rcd: $rcd,
                    ppt: $ppt,
                    pt: $pt,
                    modal_premium: $modal_premium,
                    fund_1: $fund_1,
                    fund_1_alloc: $fund_1_alloc,
                    fund_2: $fund_2,
                    fund_2_alloc: $fund_2_alloc,
                    fund_3: $fund_3,
                    fund_3_alloc: $fund_3_alloc,
                    fund_4: $fund_4,
                    fund_4_alloc: $fund_4_alloc,
                    fund_5: $fund_5,
                    fund_5_alloc: $fund_5_alloc,
                    fund_6: $fund_6,
                    fund_6_alloc: $fund_6_alloc,
                    fund_7: $fund_7,
                    fund_7_alloc: $fund_7_alloc,
                    premium_frequency: $premium_frequency,
                    gender: $gender,
                    fund_value: $fund_value,
                    ptd: $ptd,
                    dob: $dob,
                    sum_assured: $sum_assured,
                    valuation_date: $valuation_date,
                    policy_number: $policy_number,
                    module:$module,
                    level_cover_period:$level_cover_period,
                    max_projection_year:$max_projection_year,
                    is_pdf:$is_pdf,
                    calc_pay_zero:$calc_pay_zero,
                    calc_pay_one:$calc_pay_one,
                    calc_pay_three:$calc_pay_three,
                    policy_holder_name:$policy_holder_name
            ) {                   
                    valuation_date
                    maturity_date
                    rcd
                    ppt
                    pt
                    growth_rate,
                    premium_summary {
                        summary_type
                        yearly_value {
                            year
                            total_investment
                            surrender_bank_return
                            fund_value_8
                            fund_value_4
                            historic_value
                            bond_fund_historical
                        }
                            fv_end_8
                            fv_end_4
                            fv_end_historic
                            irr_8  
                            irr_4  
                            irr_historic
                            effective_charges_rest_period_8
                            effective_charges_rest_period_4
                            effective_charges_rest_period_historic
                        
                    }
                    uin
                    policy_number
                    sum_assured
                    fund_value
                    effective_returns_till_date
                    effective_charges_till_date
                    current_surrender_value
                    total_ppt_premium
                    premium_paid_till_date
                    premium_mode
                    product_name
            }
            }
                '''

        # Define the variables
        variables = {

            "uin": uin,
                "plan_option": plan_option,
                "pw": pw,
                "rcd": rcd,
                "ppt": ppt,
                "pt": pt,
                "modal_premium": modal_premium,
                "fund_1": fund_1,
                "fund_1_alloc": fund_1_alloc,
                "fund_2": fund_2,
                "fund_2_alloc": fund_2_alloc,
                "fund_3": fund_3,
                "fund_3_alloc": fund_3_alloc,
                "fund_4": fund_4,
                "fund_4_alloc": fund_4_alloc,
                "fund_5": fund_5,
                "fund_5_alloc": fund_5_alloc,
                "fund_6": fund_6,
                "fund_6_alloc": fund_6_alloc,
                "fund_7": fund_7,
                "fund_7_alloc": fund_7_alloc,
                "premium_frequency": premium_frequency,
                "gender": gender,
                "fund_value": fund_value,
                "ptd": ptd,
                "dob": dob,
                "sum_assured": sum_assured,
                "valuation_date": valuation_date,
                "policy_number": policy_number,
                "module": "bajaj",
                "level_cover_period": level_cover_period,
                "max_projection_year":pt,
                "is_pdf":True,
                "calc_pay_zero":False,
                "calc_pay_one":False,
                "calc_pay_three":False,
                "policy_holder_name":policy_holder_name

        }

        # Create the payload dictionary
        payload = {
            "query": query,
            "variables": variables
        }
        
        headers = { 
        'x-correlation-id': 'asdfvgb',
        'Content-Type': 'application/json',
        'authorizationToken': 'fixed_UUID'
        }

        payload = json.dumps(payload)
        response = requests.request("POST", url, headers=headers, data=payload)
        
    # Process the API response and collect the calculated result
        result = response.json()

        try:
            if result["data"]['Projection'] is not None:
                print(sr)
                excel_headers = ""
                python_output = result["data"]['Projection']
                premium_summary = result["data"]["Projection"]["premium_summary"]

                extracted_data = {
                    "Sr no": sr,
                    "Policy number": python_output['policy_number'],
                    "Product Name": python_output['product_name'],
                    "Premium Frequency": python_output['premium_mode'],
                    "Policy Holder Name": policy_holder_name,
                    "UIN": python_output['uin'],
                    "DOB": dob,
                    "Gender": gender,                    
                    "Growth Rate": python_output['growth_rate'],
                    "Policy Term": python_output['pt'],
                    "PTD": ptd,
                    "PPT": ppt,
                    "RCD": rcd,
                    "Sum Assured": sum_assured,
                    "Premium Paid Till date":python_output['premium_paid_till_date'],
                    "Calculation Date":valuation_date,
                    "Current Surrender Value":fund_value,
                    "Effective Returns":python_output['effective_returns_till_date'],
                    # "Effective Charges":python_output['effective_charges_till_date'],
                    "Total Premium To Be Paid":python_output['total_ppt_premium'],
                    "Maturity Date":python_output['maturity_date'],
                    "Premium End Date": (datetime.strptime(python_output['rcd'], "%d-%m-%Y") + relativedelta(years=python_output['ppt'])).strftime("%d-%m-%Y")                }
                for summary in range(1):
                    if premium_summary[summary]['summary_type'] == "Pay All":
                        extracted_data.update({
                            # "Fund Value @8%":premium_summary[summary]['fv_end_8'],
                            # "Fund Value @4%":premium_summary[summary]['fv_end_4'],
                            "Fund Value @Hist":premium_summary[summary]['fv_end_historic'],
                            # "IRR @ 8%":premium_summary[summary]['irr_8'],
                            # "IRR @ 4%":premium_summary[summary]['irr_4'],
                            # "IRR @ Hist":premium_summary[summary]['irr_historic'],
                            # "Effective Charges Rest @8":premium_summary[summary]['effective_charges_rest_period_8'],
                            # "Effective Charges Rest @4":premium_summary[summary]['effective_charges_rest_period_4'],
                            # "Effective Charges Rest @Hist":premium_summary[summary]['effective_charges_rest_period_historic'],

                        })
                full_python_output.append(extracted_data)
            else:
                print(json.dumps(variables))
                dict_not_run_case = row.to_dict()
                list_bajaj_not_run_cases.append(dict_not_run_case)
                errors = response.json()['errors']
                print(errors)     
        except Exception as e:
            dict_not_run_case = row.to_dict()
            list_bajaj_not_run_cases.append(dict_not_run_case)
            errors = response.json()['errors']
            print(json.dumps(variables))
            print(errors)
        
df_python_output    = pd.DataFrame(full_python_output)

ul_comparison_output = 'assets/output/output_{}.xlsx'.format(policy_upper_range)
with pd.ExcelWriter(ul_comparison_output) as writer:
    
    df_python_output.to_excel(writer, sheet_name='python_op', index=False)


#not run cases
df_not_run_cases    = pd.DataFrame(list_bajaj_not_run_cases)
ul_not_run_cases = 'assets/output/not_run_cases/remaining{}.xlsx'.format(policy_upper_range)
with pd.ExcelWriter(ul_not_run_cases) as writer: 
    df_not_run_cases.to_excel(writer, sheet_name='Sheet1', index=False, header=True)






