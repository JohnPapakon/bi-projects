import pandas as pd
import numpy as np
from random import choice, randint

# Ορισμός ημερομηνιών έναρξης και λήξης
start_date = pd.to_datetime('2023-01-01')
end_date = pd.to_datetime('2025-11-07')

# Κατηγορίες εξόδων και εσόδων
expense_categories = [
    'delivery', 'σινεμα', 'φαγητο & ποτό', 'supermarket', 'συναυλίες',
    'iris', 'κούρεμα', 'καφές', 'μετακίνηση', 'αποταμίευση'
]
income_categories = ['iris φίλων', 'επενδύσεις', 'μισθοδοσία']

# Δημιουργία εικονικών εξόδων
n_expenses = 300
expense_dates = pd.to_datetime(np.random.choice(
    pd.date_range(start_date, end_date), n_expenses, replace=False
))
expense_data = {
    'Ημερομηνία': expense_dates.sort_values().strftime('%d/%m/%Y'),  # μόνο ημερομηνία (ευρωπαϊκή μορφή)
    'Κατηγορία Εξόδων': np.random.choice(expense_categories, n_expenses),
    'Έξοδα': [randint(3, 100) for _ in range(n_expenses)]
}
df_expenses = pd.DataFrame(expense_data)

# Δημιουργία εικονικών εσόδων
months = pd.date_range(start_date, end_date, freq='MS')
income_dates = []
income_types = []
income_amounts = []

for date in months:
    # Μισθοδοσία μία φορά το μήνα, πάντα 1400€
    income_dates.append(date)
    income_types.append("μισθοδοσία")
    income_amounts.append(1400)
    # Τυχαίες επιπλέον καταχωρήσεις κάθε μήνα
    for _ in range(randint(0,2)):
        income_dates.append(date + pd.Timedelta(days=randint(1,25)))
        cat = choice([c for c in income_categories if c != 'μισθοδοσία'])
        income_types.append(cat)
        income_amounts.append(randint(20,300))

df_incomes = pd.DataFrame({
    'Ημερομηνία': pd.Series(income_dates).sort_values().dt.strftime('%d/%m/%Y'),  # μόνο ημερομηνία (ευρωπαϊκή μορφή)
    'Κατηγορία Εσόδων': income_types,
    'Έσοδα': income_amounts
})

# Ταξινόμηση & reset index
df_expenses = df_expenses.reset_index(drop=True)
df_incomes = df_incomes.reset_index(drop=True)

# Εξαγωγή στο ίδιο Excel αρχείο, αριστερά έξοδα, δεξιά έσοδα
file_path = 'dummy_budget_eu.xlsx'
with pd.ExcelWriter(file_path) as writer:
    df_expenses.to_excel(writer, index=False, startrow=0, startcol=0, sheet_name='Sheet1')
    df_incomes.to_excel(writer, index=False, startrow=0, startcol=7, sheet_name='Sheet1')
print(f'Το αρχείο δημιουργήθηκε: {file_path}')

