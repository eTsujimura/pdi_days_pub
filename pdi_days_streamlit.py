import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from matplotlib import pyplot

# Function to calculate workday without using workdays module
def calculate_workday(start_date, days, holidays):
    current_date = start_date
    while days > 0:
        current_date += datetime.timedelta(days=1)
        if current_date.weekday() != 6 and current_date not in holidays:  # Skip Sundays and holidays
            days -= 1
    return current_date

# Function to complete days after PDI in
def complete_days_after_PDI_in(sheet_pdiDays, model, qty):
    if qty == 0:
        return [0]

    try:
        d_model = sheet_pdiDays[model]
    except:
        model = str(model)
        d_model = sheet_pdiDays[model]

    max_day = d_model.index.max()
    cmp_list = []
    for d in range(max_day):
        try:
            cmp = d_model[d] * qty
            cmp = round(cmp)
            cmp_list.append(cmp)
        except:
            cmp_list.append(0)

        if sum(cmp_list) >= qty:
            break

    if sum(cmp_list) != qty:
        diff = qty - sum(cmp_list)
        peak = cmp_list.index(max(cmp_list))
        cmp_list[peak] = cmp_list[peak] + diff

    while cmp_list[-1] == 0:
        cmp_list.pop()

    return cmp_list

# Streamlit app
st.title("PDI-Out Forecast")

# Upload TTL_PLAN_input.xlsx
ttl_plan_file = st.file_uploader("Upload TTL_PLAN_input.xlsx", type=["xlsx"])
# Upload PDI_DAYS2.xlsx
pdi_days_file = st.file_uploader("Upload PDI_DAYS2.xlsx", type=["xlsx"])

if ttl_plan_file and pdi_days_file:
    # Read the uploaded files
    sheet_day1 = pd.read_excel(ttl_plan_file, sheet_name='INPUT', engine='openpyxl')
    sheet_pdiDays = pd.read_excel(pdi_days_file, sheet_name='INPUT', index_col=0, skiprows=[0], skipfooter=1, engine='openpyxl')
    sheet_schedule = pd.read_excel(ttl_plan_file, sheet_name='INPUT', index_col=0, skiprows=[0, 1], engine='openpyxl')

    # Set day 1
    day1 = sheet_day1.iloc[0, 1].date()

    # Additional holidays
    addtionalHolydays = [
        datetime.date(2024, 12, 7),
        datetime.date(2024, 12, 14),
        datetime.date(2024, 12, 29),
        datetime.date(2024, 12, 30),
        datetime.date(2024, 12, 31),
        datetime.date(2025, 1, 1),
        datetime.date(2025, 1, 2),
        datetime.date(2025, 1, 3),
        datetime.date(2025, 1, 5),
        datetime.date(2025, 5, 5),
        datetime.date(2025, 5, 6),
        datetime.date(2025, 5, 10),
        datetime.date(2025, 5, 17),
        datetime.date(2025, 5, 24),
        datetime.date(2025, 6, 6),
        datetime.date(2025, 6, 7),
        datetime.date(2025, 6, 21),
    ]

    models = sheet_schedule.index
    days = sheet_schedule.columns

    data = dict()
    for model in models:
        d = dict()
        tmp_car_cnt = 0
        for day in days:
            tmpDay = day1 + datetime.timedelta(days=day-1)
            pdi_in = sheet_schedule[day][model]
            pdi_in = pdi_in.item()
            try:
                pdi_in = int(pdi_in)
            except:
                pdi_in = 0

            tmp_car_cnt += pdi_in
            pdi_outs = complete_days_after_PDI_in(sheet_pdiDays, model, pdi_in)

            cnt = 0
            for out in pdi_outs:
                tmpTmpDay = calculate_workday(tmpDay, cnt, addtionalHolydays)
                if out == 0:
                    pass
                elif tmpTmpDay in d.keys():
                    curr = d[tmpTmpDay]
                    d[tmpTmpDay] = curr + out
                else:
                    d[tmpTmpDay] = out
                cnt += 1

        data[model] = d

    df = pd.DataFrame(data)
    sorted_dates = df.index.sort_values()
    lastDate = day1

    if sorted_dates[0] != day1:
        df.loc[day1] = 0

    for d in sorted_dates:
        diff = (d - lastDate).days
        if diff > 1:
            for dd in range(diff - 1):
                addDay = lastDate + datetime.timedelta(dd + 1)
                df.loc[addDay] = 0
        lastDate = d

    df.fillna(0, inplace=True)
    df = df.sort_index()
    df = df.astype('int')

    # Write to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True, sheet_name='PDI_OUT')
    output.seek(0)

    # Display DataFrame
    st.write(df)

    # Download button
    st.download_button(
        label="Download PDI_OUT.xlsx",
        data=output,
        file_name="PDI_OUT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Plotting
    ttl_pdi_out = df.T.sum()
    x = df.index
    fig = pyplot.figure(figsize=(8, 5))
    val_out = list(df.T.sum().values)
    val_out = [i.item() for i in val_out]
    val_in = list(sheet_schedule.sum().values)
    val_in = [int(i.item()) for i in val_in]

    if len(val_out) > len(val_in):
        val_in = val_in + (len(val_out) - len(val_in)) * [0]
    elif len(val_out) < len(val_in):
        val_in = val_in[:len(val_out)]

    block = 0
    val_block = []
    for a, b in zip(val_in, val_out):
        block += a - b
        val_block.append(block)

    pyplot.plot(x, val_in, label="PDI_In")
    pyplot.plot(x, val_out, label="PDI_Out")
    pyplot.plot(x, val_block, label="In PDI")

    pyplot.legend()
    pyplot.ylabel('qty')
    fig.autofmt_xdate()
    st.pyplot(fig)

