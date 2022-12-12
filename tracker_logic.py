import pandas as pd
import datetime as dt
from os .path import exists

def format_excel():
    df = pd.read_csv('UMCC Tracker.csv')
    tmr_excel = pd.ExcelWriter('UMCC Tracker.xlsx', engine='xlsxwriter')
    df.to_excel(tmr_excel, index=False, sheet_name='Movements', startrow=2)

    workbook = tmr_excel.book
    worksheet = tmr_excel.sheets['Movements']

    format_small = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'border': 1})
    format_small.set_font_size(12)

    #format_big = workbook.add_format({'text_wrap': True, 'align': 'left', 'valign': 'vcenter', 'border': 1})
    #format_big.set_font_size(12)

    update_time = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    update_time.set_font_size(10)

    title_format = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#90E0EF'})
    title_format.set_font_size(30)

    color_one = workbook.add_format({'border': 1, 'valign': 'vcenter', 'bg_color': '#CAF0F8', 'text_wrap': True})
    color_one.set_align('vcenter')

    for ind, i in enumerate(df.values):
        if ind % 2 == 0:
            worksheet.set_row(ind + 2, None, color_one)
    
    worksheet.merge_range('A1:J1', f'UMCC Tracker', title_format)
    worksheet.merge_range('A2:J2', f'Updated {dt.datetime.now()}', update_time)

    worksheet.set_column('A:B', 20, format_small)

    worksheet.set_column('D:D', 15, format_small)
    worksheet.set_column('F:F', 15, format_small)

    worksheet.set_column('C:C', 30, format_small)
    worksheet.set_column('E:E', 30, format_small)
    worksheet.set_column('G:G', 30, format_small)
    worksheet.set_column('J:J', 40, format_small)

    worksheet.set_column('H:I', 20, format_small)

    tmr_excel.close()

    return tmr_excel

def export_movement_tracker(data):
    print(f"\nCREATING UMCC TRACKER @ {dt.datetime.now()}\n")
    df_tmrs = pd.DataFrame.from_dict(data, orient='index')
    df_tmrs.columns = ['Tag', 'Num', 'Support Requested', 'Start Date/Time', 'Pickup Location', 'End Date/Time',
                        'Drop Off Location', 'Supporting Unit', 'Status', 'Additional Comments']
    df_tmrs.to_csv('UMCC Tracker.csv', index=False)

def import_tmrs(tcpt_file):
    imported_tmrs = {}
    # reading excel
    df1 = pd.read_excel(tcpt_file)

    # creating a temporary list to store data
    temp_list = [v for v in df1.values]
    temp_list.pop(0)
    temp_list.pop()

    # adding to dictonary for better organization & cleaning up the information
    for index, i in enumerate(temp_list):
        name = i[9]
        num = i[1]
        start_dt = i[3]
        end_dt = i[5]
        pu_location = i[2]
        do_location = i[4]
        support_unit = i[8]
        status = i[10]

        # breaking down date and time into seperate variables
        s_year, s_date, s_time = start_dt[6:10], start_dt[:10], start_dt[11:16]
        e_year, e_date, e_time = end_dt[6:10], end_dt[:10], end_dt[11:16]

        # start date conversion
        a_s, b_s = int(start_dt[6:10]), int(start_dt[:2])
        start_month = dt.datetime(a_s, b_s, 1).strftime("%B")
        s_date = f'{s_date[3:5]} {start_month[:3]} {s_year}'
        # end date conversion
        a_e, b_e = int(end_dt[6:10]), int(end_dt[:2])
        end_month = dt.datetime(a_e, b_e, 1).strftime("%B")
        e_date = f'{e_date[3:5]} {end_month[:3]} {e_year}'

        imported_tmrs[index] = [name, num, '', f'{s_time} {s_date}', pu_location, f'{e_time} {e_date}',
        do_location, support_unit, status, '']

    return imported_tmrs

def import_movements(imported_tmr_dict):
    final_tmr_dict = {}
    existing_tmrs = {}
    if exists('UMCC Tracker.csv'):
        # checking if the tracker already exits
        df = pd.read_csv('UMCC Tracker.csv')

        # creating a temporary list to store data
        temp_list2 = [v for v in df.values]

        # adding to dictonary for better organization & cleaning up the information
        for index, i in enumerate(temp_list2):
            existing_tmrs[index+1] = {'tmr name': i[0], 'tmr num': i[1], 'support needed': i[2],
            'start dtg': i[3], 'pickup location': i[4], 'end dtg': i[5], 'dropoff location': i[6],
            'support unit': i[7], 'status': i[8], 'comments': i[9]}
        
        existing_list = []
        for i in existing_tmrs.values():
            existing_list.append(i['tmr num'])

        for ind, n in enumerate(imported_tmr_dict.values()):
            if n['tmr num'] in existing_list:
                tmr_info = existing_list.index(n['tmr num'])+1
                name = n['tmr name']
                sd = n['start dtg']
                pu = n['pickup location']
                ed = n['end dtg']
                do = n['dropoff location']
                su = n['support unit']
                sa = n['status']
                cm = existing_tmrs[tmr_info]['comments']
                sn = existing_tmrs[tmr_info]['support needed']

                final_tmr_dict[ind] = {
                    'tmr name': name,
                    'tmr num': n['tmr num'],
                    'support needed': sn,
                    'start dtg': sd,
                    'pickup location': pu,
                    'end dtg': ed,
                    'dropoff location': do,
                    'support unit': su,
                    'status': sa,
                    'comments': cm
                }
            else:
                print(f"New TMR: {n['tmr num']}\n")
                
                # support_needed = input(f"Enter support submitted for {n['tmr name']} {n['tmr num']}, from {n['start dtg']} -> {n['end dtg']}: ")
                # comments = input(f"Enter any additional comments for {n['tmr name']} {n['tmr num']}: ")
                # print('\n')

                final_tmr_dict[ind] = {
                    'tmr name': n['tmr name'],
                    'tmr num': n['tmr num'],
                    'support needed': ' ',
                    'start dtg': n['start dtg'],
                    'pickup location': n['pickup location'],
                    'end dtg': n['end dtg'],
                    'dropoff location': n['dropoff location'],
                    'support unit': n['support unit'],
                    'status': n['status'],
                    'comments': ' '
                }
    else:
        for ind, n in enumerate(imported_tmr_dict.values()):
            print(f"new tmr {n}")
            name = n['tmr name']
            sd = n['start dtg']
            pu = n['pickup location']
            ed = n['end dtg']
            do = n['dropoff location']
            su = n['support unit']
            sa = n['status']

            final_tmr_dict[ind] = {
                'tmr name': name,
                'tmr num': n['tmr num'],
                'support needed': ' ',
                'start dtg': sd,
                'pickup location': pu,
                'end dtg': ed,
                'dropoff location': do,
                'support unit': su,
                'status': sa,
                'comments': ' '
            }
    if len(final_tmr_dict) > 0:
        return final_tmr_dict
    else:
        return 'NO CURRENT TMRS SUBMITTED'

def remove_old(tmr_dic):
    current_list = {}

    for ind, v in enumerate(tmr_dic.values()):
        day = v[5][6:8]
        month = v[5][9:12]
        year = v[5][13:]

        tmr_date = f"{day}/{month}/{year}"

        d = dt.datetime.strptime(tmr_date, '%d/%b/%Y')
        today = dt.datetime.today()

        if d <= today:
            print(f"{v[0]} {v[1]}: is old\n")
        else:
            print(f"{v[0]} {v[1]}: is current\n")
            # add to the dictonary of current tmr

            name = v[0]
            sn = v[2]
            sd = v[3]
            pu = v[4]
            ed = v[5]
            do = v[6]
            su = v[7]
            sa = v[8]
            cm = v[9]

            current_list[ind] = {
                'tmr name': name,
                'tmr num': v[1],
                'support needed': sn,
                'start dtg': sd,
                'pickup location': pu,
                'end dtg': ed,
                'dropoff location': do,
                'support unit': su,
                'status': sa,
                'comments': cm
            }

    return current_list