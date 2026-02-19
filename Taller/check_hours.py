import pandas as pd

# Load data
data_rows = []
with open('/Users/javierfernandez/Desktop/Taller/Libro1.csv', 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

# Parse header
header_line = lines[0].strip()
if header_line.startswith('"') and header_line.endswith('"'):
    header = header_line[1:-1].split(';')

# Parse data lines
for line in lines[1:]:
    line = line.strip()
    if line.startswith('"') and line.endswith('"'):
        row_data = line[1:-1].split(';')
        data_rows.append(row_data)

df = pd.DataFrame(data_rows, columns=header)

# Convert Time column
df['Time'] = df['Time'].str.replace(',', '.').astype(float)
df['Hour'] = (df['Time'] * 24).round().astype(int).clip(0, 23)

print("Análisis de horas de actividad:")
print(f"Hora mínima: {df['Hour'].min()}:00")
print(f"Hora máxima: {df['Hour'].max()}:00")
print(f"Rango de actividad: {df['Hour'].min()}:00 - {df['Hour'].max()}:00")

print("\nDistribución por hora:")
hourly_count = df.groupby('Hour').size().sort_index()
for hour, count in hourly_count.items():
    print(f"{hour:2d}:00 - {count:3d} transacciones")

print(f"\nTotal de transacciones: {len(df)}")
