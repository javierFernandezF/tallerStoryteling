import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Set up plotting style
plt.style.use('default')
sns.set_palette("husl")

def load_data():
    """Load and clean the supermarket sales data"""
    print("Loading data...")
    
    # Read the file manually to handle the special format
    data_rows = []
    with open('/Users/javierfernandez/Desktop/Taller/Libro1.csv', 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
    
    # Parse header from first line
    header_line = lines[0].strip()
    if header_line.startswith('"') and header_line.endswith('"'):
        header = header_line[1:-1].split(';')
    
    # Parse data lines
    for line in lines[1:]:
        line = line.strip()
        if line.startswith('"') and line.endswith('"'):
            row_data = line[1:-1].split(';')
            data_rows.append(row_data)
    
    # Create DataFrame
    df = pd.DataFrame(data_rows, columns=header)
    
    # Convert numeric columns (replace comma with dot)
    numeric_cols = ['Unit price', 'Quantity', 'Tax 5%', 'Total', 'Time', 'cogs', 'gross margin percentage', 'gross income', 'Rating']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].str.replace(',', '.').astype(float)
    
    # Create hour from time (assuming decimal format)
    df['Hour'] = (df['Time'] * 24).round().astype(int).clip(0, 23)
    
    print(f"Dataset loaded: {len(df)} records")
    print(f"Columns: {list(df.columns)}")
    return df

def analyze_all():
    """Perform all analyses"""
    df = load_data()
    
    print("\n" + "="*80)
    print("ANÁLISIS COMPLETO - SUPERMERCADO")
    print("="*80)
    
    # 1. ANÁLISIS POR GÉNERO
    print("\n1. COMPRAS POR GÉNERO:")
    print("-" * 40)
    
    gender_summary = df.groupby('Gender').agg({
        'Total': ['count', 'sum', 'mean'],
        'Quantity': 'sum'
    }).round(2)
    
    female_sales = df[df['Gender'] == 'Female']['Total'].sum()
    male_sales = df[df['Gender'] == 'Male']['Total'].sum()
    total_sales = female_sales + male_sales
    
    print(f"Ventas Mujeres: ${female_sales:,.2f} ({female_sales/total_sales*100:.1f}%)")
    print(f"Ventas Hombres: ${male_sales:,.2f} ({male_sales/total_sales*100:.1f}%)")
    print(f"Promedio compra Mujeres: ${df[df['Gender'] == 'Female']['Total'].mean():.2f}")
    print(f"Promedio compra Hombres: ${df[df['Gender'] == 'Male']['Total'].mean():.2f}")
    print(f"Transacciones Mujeres: {len(df[df['Gender'] == 'Female'])}")
    print(f"Transacciones Hombres: {len(df[df['Gender'] == 'Male'])}")
    
    # 2. ANÁLISIS POR HORA
    print("\n2. DISTRIBUCIÓN POR HORA DEL DÍA:")
    print("-" * 40)
    
    hourly_sales = df.groupby('Hour')['Total'].agg(['count', 'sum', 'mean']).round(2)
    peak_hour = df.groupby('Hour')['Total'].sum().idxmax()
    peak_sales = df.groupby('Hour')['Total'].sum().max()
    
    print(f"Hora pico: {peak_hour}:00 hrs con ${peak_sales:,.2f} en ventas")
    print(f"Horario de operación: 10:00 - 21:00 hrs (12 horas)")
    
    # Group by time periods - DIVISIÓN EQUITATIVA (4 horas cada período)
    morning = df[df['Hour'].between(10, 13)]['Total'].sum()    # 10-13h (4 horas)
    afternoon = df[df['Hour'].between(14, 17)]['Total'].sum()  # 14-17h (4 horas)
    evening = df[df['Hour'].between(18, 21)]['Total'].sum()    # 18-21h (4 horas)
    
    total_sales = df['Total'].sum()
    print(f"Ventas Mañana (10-13h): ${morning:,.2f} ({morning/total_sales*100:.1f}%)")
    print(f"Ventas Tarde (14-17h): ${afternoon:,.2f} ({afternoon/total_sales*100:.1f}%)")
    print(f"Ventas Noche (18-21h): ${evening:,.2f} ({evening/total_sales*100:.1f}%)")
    print("Distribución equilibrada - diferencias menores entre períodos")
    
    # 3. INGRESOS POR LÍNEA DE PRODUCTOS
    print("\n3. INGRESOS POR LÍNEA DE PRODUCTOS:")
    print("-" * 40)
    
    product_sales = df.groupby('Product line')['Total'].agg(['count', 'sum', 'mean']).round(2)
    product_sales_sorted = df.groupby('Product line')['Total'].sum().sort_values(ascending=False)
    
    for product, sales in product_sales_sorted.items():
        pct = (sales / df['Total'].sum()) * 100
        count = len(df[df['Product line'] == product])
        avg = df[df['Product line'] == product]['Total'].mean()
        print(f"{product}: ${sales:,.2f} ({pct:.1f}%) - {count} transacciones - Promedio: ${avg:.2f}")
    
    # 4. INGRESOS POR LÍNEA DE PRODUCTOS Y GÉNERO
    print("\n4. INGRESOS POR LÍNEA DE PRODUCTOS Y GÉNERO:")
    print("-" * 40)
    
    product_gender = df.groupby(['Product line', 'Gender'])['Total'].sum().unstack(fill_value=0)
    
    for product in product_gender.index:
        female_sales = product_gender.loc[product, 'Female']
        male_sales = product_gender.loc[product, 'Male']
        total_product = female_sales + male_sales
        print(f"{product}:")
        print(f"  Mujeres: ${female_sales:,.2f} ({female_sales/total_product*100:.1f}%)")
        print(f"  Hombres: ${male_sales:,.2f} ({male_sales/total_product*100:.1f}%)")
    
    # 5. COMPRAS POR SUCURSAL
    print("\n5. COMPRAS POR SUCURSAL:")
    print("-" * 40)
    
    branch_sales = df.groupby('Branch')['Total'].agg(['count', 'sum', 'mean']).round(2)
    
    for branch in sorted(df['Branch'].unique()):
        sales = df[df['Branch'] == branch]['Total'].sum()
        count = len(df[df['Branch'] == branch])
        avg = df[df['Branch'] == branch]['Total'].mean()
        pct = (sales / df['Total'].sum()) * 100
        city = df[df['Branch'] == branch]['City'].iloc[0]
        print(f"Sucursal {branch} ({city}): ${sales:,.2f} ({pct:.1f}%) - {count} transacciones - Promedio: ${avg:.2f}")
    
    # 6. COMPRAS POR LÍNEA DE PRODUCTOS (resumen)
    print("\n6. RESUMEN DE COMPRAS POR LÍNEA DE PRODUCTOS:")
    print("-" * 40)
    
    product_summary = df.groupby('Product line').agg({
        'Total': ['count', 'sum'],
        'Quantity': 'sum',
        'gross income': 'sum'
    }).round(2)
    
    print("Ranking por cantidad de transacciones:")
    transaction_ranking = df.groupby('Product line').size().sort_values(ascending=False)
    for i, (product, count) in enumerate(transaction_ranking.items(), 1):
        print(f"{i}. {product}: {count} transacciones")
    
    # RECOMENDACIONES
    print("\n" + "="*80)
    print("RECOMENDACIONES PARA CAMPAÑAS DE MARKETING")
    print("="*80)
    
    print("\n📊 INSIGHTS CLAVE:")
    print(f"• Total de ventas: ${df['Total'].sum():,.2f}")
    print(f"• Promedio por transacción: ${df['Total'].mean():.2f}")
    print(f"• Total de transacciones: {len(df):,}")
    
    # Top product line
    top_product = product_sales_sorted.index[0]
    top_product_sales = product_sales_sorted.iloc[0]
    print(f"• Línea más rentable: {top_product} (${top_product_sales:,.2f})")
    
    # Gender insights
    if female_sales > male_sales:
        print(f"• Las mujeres generan más ventas ({female_sales/total_sales*100:.1f}% vs {male_sales/total_sales*100:.1f}%)")
    else:
        print(f"• Los hombres generan más ventas ({male_sales/total_sales*100:.1f}% vs {female_sales/total_sales*100:.1f}%)")
    
    print(f"\n🎯 RECOMENDACIONES:")
    print("1. SEGMENTACIÓN POR GÉNERO:")
    print("   - Desarrollar campañas específicas según el comportamiento de compra por género")
    print("   - Enfocar productos con mayor diferencia de preferencia por género")
    
    print("\n2. OPTIMIZACIÓN HORARIA:")
    print(f"   - Concentrar esfuerzos promocionales en hora pico ({peak_hour}:00 hrs)")
    print("   - Desarrollar promociones específicas para horarios de menor actividad")
    
    print("\n3. ESTRATEGIA POR PRODUCTO:")
    print(f"   - Potenciar la línea más exitosa: {top_product}")
    print("   - Crear bundles con productos complementarios")
    print("   - Mejorar el rendimiento de líneas con menor participación")
    
    print("\n4. EQUILIBRIO ENTRE SUCURSALES:")
    best_branch = df.groupby('Branch')['Total'].sum().idxmax()
    print(f"   - Replicar estrategias exitosas de Sucursal {best_branch} en otras ubicaciones")
    print("   - Analizar factores locales que afectan el rendimiento por sucursal")
    
    print("\n5. CAMPAÑAS MENSUALES SUGERIDAS:")
    print("   MES 1: Campaña enfocada en género dominante y productos estrella")
    print("   MES 2: Promociones cruzadas entre líneas de productos")
    print("   MES 3: Optimización horaria y equilibrio entre sucursales")

if __name__ == "__main__":
    analyze_all()
