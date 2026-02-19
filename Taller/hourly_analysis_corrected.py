import pandas as pd

# Load data
data_rows = []
with open('/Users/javierfernandez/Desktop/Taller/Libro1.csv', 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

header_line = lines[0].strip()
if header_line.startswith('"') and header_line.endswith('"'):
    header = header_line[1:-1].split(';')

for line in lines[1:]:
    line = line.strip()
    if line.startswith('"') and line.endswith('"'):
        row_data = line[1:-1].split(';')
        data_rows.append(row_data)

df = pd.DataFrame(data_rows, columns=header)

# Convert numeric columns
df['Total'] = df['Total'].str.replace(',', '.').astype(float)
df['Time'] = df['Time'].str.replace(',', '.').astype(float)
df['Hour'] = (df['Time'] * 24).round().astype(int).clip(0, 23)

print("ANÁLISIS HORARIO CORREGIDO - DIVISIÓN EQUITATIVA")
print("=" * 60)

# Horario de operación: 10:00 - 21:00 (12 horas)
# División equitativa en 3 períodos de 4 horas cada uno:

morning = df[df['Hour'].between(10, 13)]['Total'].sum()    # 10-13h (4 horas)
afternoon = df[df['Hour'].between(14, 17)]['Total'].sum()  # 14-17h (4 horas)  
evening = df[df['Hour'].between(18, 21)]['Total'].sum()    # 18-21h (4 horas)

total_sales = df['Total'].sum()

print(f"Horario de operación: 10:00 - 21:00 hrs (12 horas)")
print(f"División equitativa en períodos de 4 horas:")
print()
print(f"Mañana (10-13h):  ${morning:,.2f} ({morning/total_sales*100:.1f}%)")
print(f"Tarde (14-17h):   ${afternoon:,.2f} ({afternoon/total_sales*100:.1f}%)")
print(f"Noche (18-21h):   ${evening:,.2f} ({evening/total_sales*100:.1f}%)")
print()

# Análisis detallado por hora
print("DETALLE POR HORA:")
print("-" * 30)
hourly_sales = df.groupby('Hour')['Total'].sum()
hourly_count = df.groupby('Hour').size()

for hour in range(10, 22):
    sales = hourly_sales.get(hour, 0)
    count = hourly_count.get(hour, 0)
    pct = (sales / total_sales) * 100
    print(f"{hour:2d}:00 - ${sales:8,.2f} ({pct:4.1f}%) - {count:3d} transacciones")

peak_hour = hourly_sales.idxmax()
peak_sales = hourly_sales.max()
print()
print(f"Hora pico: {peak_hour}:00 hrs con ${peak_sales:,.2f}")

# Comparación de períodos
print()
print("COMPARACIÓN DE PERÍODOS (4 horas cada uno):")
print("-" * 45)
periods = [
    ("Mañana (10-13h)", morning),
    ("Tarde (14-17h)", afternoon), 
    ("Noche (18-21h)", evening)
]

periods_sorted = sorted(periods, key=lambda x: x[1], reverse=True)
for i, (period, sales) in enumerate(periods_sorted, 1):
    pct = (sales / total_sales) * 100
    print(f"{i}. {period}: ${sales:,.2f} ({pct:.1f}%)")
