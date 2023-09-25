import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('data.xlsx', engine='openpyxl')
data['receiving_date'] = pd.to_datetime(data['receiving_date'], errors='coerce')
valid_deals = data[(data['status'] == 'ОПЛАЧЕНО') & (data['receiving_date'] <= pd.Timestamp('now'))]


july_2021_revenue = valid_deals[valid_deals['receiving_date'].dt.month == 7]['sum'].sum()
print(f'Общая выручка за июль 2021: {july_2021_revenue}')


revenue_by_month = valid_deals.groupby(valid_deals['receiving_date'].dt.to_period('M'))['sum'].sum()
revenue_by_month.plot(kind='line', title='Динамика выручки')
plt.xlabel('Месяц')
plt.ylabel('Выручка')
plt.show(block=True)


september_2021_revenue = valid_deals[valid_deals['receiving_date'].dt.month == 9].groupby('sale')['sum'].sum()
top_manager = september_2021_revenue.idxmax()
print(f'Менеджер, привлекший больше всего денег в сентябре 2021: {top_manager}')


october_2021_deals = valid_deals[valid_deals['receiving_date'].dt.month == 10]
deal_type_counts = october_2021_deals['new/current'].value_counts()
dominant_deal_type = deal_type_counts.idxmax()
print(f'Преобладающий тип сделок в октябре 2021: {dominant_deal_type}')


may_2021_deals = valid_deals[(valid_deals['receiving_date'].dt.month == 5) & (valid_deals['receiving_date'].dt.year == 2021)]
june_2021_originals = may_2021_deals[(may_2021_deals['receiving_date'].dt.month == 6) & (may_2021_deals['receiving_date'].dt.year == 2021)]
print(f'Количество оригиналов договоров по майским сделкам, полученных в июне 2021: {len(june_2021_originals)}')


def calculate_bonus(row):
    if row['new/current'] == 'новая':
        if row['status'] == 'ОПЛАЧЕНО' and row['receiving_date'].month == 7:
            return 0.07 * row['sum']
        else:
            return 0
    elif row['new/current'] == 'текущая':
        if row['status'] != 'ПРОСРОЧЕНО' and row['receiving_date'].month == 7:
            return 0.05 * row['sum'] if row['sum'] > 10000 else 0.03 * row['sum']
        else:
            return 0
    else:
        return 0

data['bonus'] = data.apply(calculate_bonus, axis=1)
manager_balances = data[data['receiving_date'].dt.month == 6].groupby('sale')['bonus'].sum()
print('Остаток каждого менеджера на 01.07.2021:')
print(manager_balances)
