import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set up plotting style
plt.style.use('default')
sns.set_palette("husl")

def load_and_clean_data(file_path):
    """Load and clean the supermarket sales data"""
    print("Loading and cleaning data...")
    
    # Read the file manually to handle the special format
    data_rows = []
    with open(file_path, 'r', encoding='utf-8-sig') as f:  # Handle BOM
        lines = f.readlines()
    
    # Parse header from first line
    header_line = lines[0].strip()
    if header_line.startswith('"') and header_line.endswith('"'):
        header = header_line[1:-1].split(';')
    else:
        header = header_line.split(';')
    
    # Parse data lines
    for line in lines[1:]:
        line = line.strip()
        if line.startswith('"') and line.endswith('"'):
            row_data = line[1:-1].split(';')
            data_rows.append(row_data)
    
    # Create DataFrame
    df = pd.DataFrame(data_rows, columns=header)
    
    # Clean column names (remove quotes if present)
    df.columns = df.columns.str.strip('"')
    
    # Convert decimal comma to decimal point for numeric columns
    numeric_columns = ['Unit price', 'Quantity', 'Tax 5%', 'Total', 'Time', 'cogs', 'gross margin percentage', 'gross income', 'Rating']
    
    for col in numeric_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '.').astype(float)
    
    # Convert Date column (appears to be in Excel serial date format)
    # Excel serial date: days since 1900-01-01
    df['Date'] = pd.to_numeric(df['Date'], errors='coerce')
    df['Date'] = pd.to_datetime(df['Date'], origin='1899-12-30', unit='D', errors='coerce')
    
    # Create hour column from Time (assuming Time is in decimal format like 0.547 = 13:07)
    df['Hour'] = (df['Time'] * 24).round().astype(int)
    df['Hour'] = df['Hour'].clip(0, 23)  # Ensure hours are between 0-23
    
    # Clean text columns
    text_columns = ['Branch', 'City', 'Customer type', 'Gender', 'Product line', 'Payment']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    
    print(f"Dataset loaded successfully with {len(df)} records and {len(df.columns)} columns")
    print(f"Date range: {df['Date'].min()} to {df['Date'].max()}")
    
    return df

def analyze_gender_purchases(df):
    """Analyze purchases by gender"""
    print("\n=== ANÁLISIS POR GÉNERO ===")
    
    # Basic statistics by gender
    gender_stats = df.groupby('Gender').agg({
        'Total': ['count', 'sum', 'mean'],
        'Quantity': ['sum', 'mean'],
        'gross income': ['sum', 'mean']
    }).round(2)
    
    print("Estadísticas por género:")
    print(gender_stats)
    
    # Create visualization
    fig, axes = plt.subplots(2, 2, figsize=(15, 10))
    fig.suptitle('Análisis de Compras por Género', fontsize=16, fontweight='bold')
    
    # Total sales by gender
    gender_totals = df.groupby('Gender')['Total'].sum()
    axes[0,0].pie(gender_totals.values, labels=gender_totals.index, autopct='%1.1f%%', startangle=90)
    axes[0,0].set_title('Distribución de Ventas Totales por Género')
    
    # Average purchase by gender
    gender_avg = df.groupby('Gender')['Total'].mean()
    axes[0,1].bar(gender_avg.index, gender_avg.values, color=['lightcoral', 'lightblue'])
    axes[0,1].set_title('Promedio de Compra por Género')
    axes[0,1].set_ylabel('Promedio ($)')
    
    # Number of transactions by gender
    gender_count = df.groupby('Gender').size()
    axes[1,0].bar(gender_count.index, gender_count.values, color=['lightcoral', 'lightblue'])
    axes[1,0].set_title('Número de Transacciones por Género')
    axes[1,0].set_ylabel('Cantidad de Transacciones')
    
    # Box plot of purchase amounts by gender
    df.boxplot(column='Total', by='Gender', ax=axes[1,1])
    axes[1,1].set_title('Distribución de Montos de Compra por Género')
    axes[1,1].set_ylabel('Monto Total ($)')
    
    plt.tight_layout()
    plt.savefig('/Users/javierfernandez/Desktop/Taller/gender_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return gender_stats

def analyze_hourly_distribution(df):
    """Analyze purchase distribution by hour"""
    print("\n=== ANÁLISIS POR HORA DEL DÍA ===")
    
    # Group by hour
    hourly_stats = df.groupby('Hour').agg({
        'Total': ['count', 'sum', 'mean']
    }).round(2)
    
    print("Estadísticas por hora:")
    print(hourly_stats.head(10))
    
    # Create visualization
    fig, axes = plt.subplots(2, 2, figsize=(15, 10))
    fig.suptitle('Análisis de Compras por Hora del Día', fontsize=16, fontweight='bold')
    
    # Number of transactions by hour
    hourly_count = df.groupby('Hour').size()
    axes[0,0].plot(hourly_count.index, hourly_count.values, marker='o', linewidth=2)
    axes[0,0].set_title('Número de Transacciones por Hora')
    axes[0,0].set_xlabel('Hora del Día')
    axes[0,0].set_ylabel('Número de Transacciones')
    axes[0,0].grid(True, alpha=0.3)
    
    # Total sales by hour
    hourly_sales = df.groupby('Hour')['Total'].sum()
    axes[0,1].bar(hourly_sales.index, hourly_sales.values, alpha=0.7)
    axes[0,1].set_title('Ventas Totales por Hora')
    axes[0,1].set_xlabel('Hora del Día')
    axes[0,1].set_ylabel('Ventas Totales ($)')
    
    # Average purchase by hour
    hourly_avg = df.groupby('Hour')['Total'].mean()
    axes[1,0].plot(hourly_avg.index, hourly_avg.values, marker='s', color='green', linewidth=2)
    axes[1,0].set_title('Promedio de Compra por Hora')
    axes[1,0].set_xlabel('Hora del Día')
    axes[1,0].set_ylabel('Promedio ($)')
    axes[1,0].grid(True, alpha=0.3)
    
    # Heatmap of transactions by hour and gender
    hourly_gender = df.groupby(['Hour', 'Gender']).size().unstack(fill_value=0)
    sns.heatmap(hourly_gender.T, annot=True, fmt='d', cmap='YlOrRd', ax=axes[1,1])
    axes[1,1].set_title('Transacciones por Hora y Género')
    axes[1,1].set_xlabel('Hora del Día')
    
    plt.tight_layout()
    plt.savefig('/Users/javierfernandez/Desktop/Taller/hourly_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return hourly_stats

def analyze_product_line_revenue(df):
    """Analyze revenue by product line"""
    print("\n=== ANÁLISIS DE INGRESOS POR LÍNEA DE PRODUCTOS ===")
    
    # Revenue by product line
    product_revenue = df.groupby('Product line').agg({
        'Total': ['count', 'sum', 'mean'],
        'gross income': ['sum', 'mean'],
        'Quantity': ['sum', 'mean']
    }).round(2)
    
    print("Ingresos por línea de productos:")
    print(product_revenue)
    
    # Create visualization
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('Análisis de Ingresos por Línea de Productos', fontsize=16, fontweight='bold')
    
    # Total revenue by product line
    revenue_by_product = df.groupby('Product line')['Total'].sum().sort_values(ascending=False)
    axes[0,0].barh(revenue_by_product.index, revenue_by_product.values)
    axes[0,0].set_title('Ingresos Totales por Línea de Productos')
    axes[0,0].set_xlabel('Ingresos Totales ($)')
    
    # Average purchase by product line
    avg_by_product = df.groupby('Product line')['Total'].mean().sort_values(ascending=False)
    axes[0,1].barh(avg_by_product.index, avg_by_product.values, color='orange')
    axes[0,1].set_title('Promedio de Compra por Línea de Productos')
    axes[0,1].set_xlabel('Promedio ($)')
    
    # Number of transactions by product line
    count_by_product = df.groupby('Product line').size().sort_values(ascending=False)
    axes[1,0].barh(count_by_product.index, count_by_product.values, color='green')
    axes[1,0].set_title('Número de Transacciones por Línea de Productos')
    axes[1,0].set_xlabel('Número de Transacciones')
    
    # Pie chart of revenue distribution
    axes[1,1].pie(revenue_by_product.values, labels=revenue_by_product.index, autopct='%1.1f%%', startangle=90)
    axes[1,1].set_title('Distribución de Ingresos por Línea de Productos')
    
    plt.tight_layout()
    plt.savefig('/Users/javierfernandez/Desktop/Taller/product_line_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return product_revenue

def analyze_product_line_by_gender(df):
    """Analyze revenue by product line and gender"""
    print("\n=== ANÁLISIS DE INGRESOS POR LÍNEA DE PRODUCTOS Y GÉNERO ===")
    
    # Revenue by product line and gender
    product_gender_revenue = df.groupby(['Product line', 'Gender']).agg({
        'Total': ['count', 'sum', 'mean'],
        'gross income': ['sum']
    }).round(2)
    
    print("Ingresos por línea de productos y género:")
    print(product_gender_revenue)
    
    # Create pivot table for visualization
    pivot_revenue = df.pivot_table(values='Total', index='Product line', columns='Gender', aggfunc='sum', fill_value=0)
    pivot_count = df.pivot_table(values='Total', index='Product line', columns='Gender', aggfunc='count', fill_value=0)
    
    # Create visualization
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('Análisis de Ingresos por Línea de Productos y Género', fontsize=16, fontweight='bold')
    
    # Heatmap of revenue by product line and gender
    sns.heatmap(pivot_revenue, annot=True, fmt='.0f', cmap='YlOrRd', ax=axes[0,0])
    axes[0,0].set_title('Ingresos Totales por Línea de Productos y Género')
    axes[0,0].set_ylabel('Línea de Productos')
    
    # Stacked bar chart
    pivot_revenue.plot(kind='bar', stacked=True, ax=axes[0,1], color=['lightcoral', 'lightblue'])
    axes[0,1].set_title('Ingresos por Línea de Productos y Género (Apilado)')
    axes[0,1].set_ylabel('Ingresos Totales ($)')
    axes[0,1].tick_params(axis='x', rotation=45)
    axes[0,1].legend(title='Género')
    
    # Grouped bar chart
    pivot_revenue.plot(kind='bar', ax=axes[1,0], color=['lightcoral', 'lightblue'])
    axes[1,0].set_title('Ingresos por Línea de Productos y Género (Agrupado)')
    axes[1,0].set_ylabel('Ingresos Totales ($)')
    axes[1,0].tick_params(axis='x', rotation=45)
    axes[1,0].legend(title='Género')
    
    # Heatmap of transaction count
    sns.heatmap(pivot_count, annot=True, fmt='d', cmap='Blues', ax=axes[1,1])
    axes[1,1].set_title('Número de Transacciones por Línea de Productos y Género')
    axes[1,1].set_ylabel('Línea de Productos')
    
    plt.tight_layout()
    plt.savefig('/Users/javierfernandez/Desktop/Taller/product_gender_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return product_gender_revenue

def analyze_branch_purchases(df):
    """Analyze purchases by branch"""
    print("\n=== ANÁLISIS DE COMPRAS POR SUCURSAL ===")
    
    # Statistics by branch
    branch_stats = df.groupby('Branch').agg({
        'Total': ['count', 'sum', 'mean'],
        'gross income': ['sum', 'mean'],
        'Quantity': ['sum', 'mean']
    }).round(2)
    
    print("Estadísticas por sucursal:")
    print(branch_stats)
    
    # Create visualization
    fig, axes = plt.subplots(2, 2, figsize=(15, 10))
    fig.suptitle('Análisis de Compras por Sucursal', fontsize=16, fontweight='bold')
    
    # Total sales by branch
    branch_sales = df.groupby('Branch')['Total'].sum()
    axes[0,0].bar(branch_sales.index, branch_sales.values, color=['red', 'green', 'blue'])
    axes[0,0].set_title('Ventas Totales por Sucursal')
    axes[0,0].set_ylabel('Ventas Totales ($)')
    
    # Number of transactions by branch
    branch_count = df.groupby('Branch').size()
    axes[0,1].bar(branch_count.index, branch_count.values, color=['red', 'green', 'blue'])
    axes[0,1].set_title('Número de Transacciones por Sucursal')
    axes[0,1].set_ylabel('Número de Transacciones')
    
    # Average purchase by branch
    branch_avg = df.groupby('Branch')['Total'].mean()
    axes[1,0].bar(branch_avg.index, branch_avg.values, color=['red', 'green', 'blue'])
    axes[1,0].set_title('Promedio de Compra por Sucursal')
    axes[1,0].set_ylabel('Promedio ($)')
    
    # Branch performance by product line
    branch_product = df.groupby(['Branch', 'Product line'])['Total'].sum().unstack(fill_value=0)
    branch_product.plot(kind='bar', stacked=True, ax=axes[1,1])
    axes[1,1].set_title('Ventas por Sucursal y Línea de Productos')
    axes[1,1].set_ylabel('Ventas Totales ($)')
    axes[1,1].tick_params(axis='x', rotation=0)
    axes[1,1].legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    
    plt.tight_layout()
    plt.savefig('/Users/javierfernandez/Desktop/Taller/branch_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return branch_stats

def generate_summary_report(df, gender_stats, hourly_stats, product_revenue, product_gender_revenue, branch_stats):
    """Generate a comprehensive summary report"""
    print("\n" + "="*80)
    print("REPORTE EJECUTIVO - ANÁLISIS DE VENTAS SUPERMERCADO")
    print("="*80)
    
    print(f"\nRESUMEN GENERAL:")
    print(f"- Total de transacciones: {len(df):,}")
    print(f"- Ingresos totales: ${df['Total'].sum():,.2f}")
    print(f"- Promedio por transacción: ${df['Total'].mean():.2f}")
    print(f"- Período analizado: {df['Date'].min().strftime('%d/%m/%Y')} - {df['Date'].max().strftime('%d/%m/%Y')}")
    
    print(f"\n1. ANÁLISIS POR GÉNERO:")
    female_total = df[df['Gender'] == 'Female']['Total'].sum()
    male_total = df[df['Gender'] == 'Male']['Total'].sum()
    female_pct = (female_total / (female_total + male_total)) * 100
    male_pct = (male_total / (female_total + male_total)) * 100
    
    print(f"- Mujeres: ${female_total:,.2f} ({female_pct:.1f}% del total)")
    print(f"- Hombres: ${male_total:,.2f} ({male_pct:.1f}% del total)")
    print(f"- Promedio compra mujeres: ${df[df['Gender'] == 'Female']['Total'].mean():.2f}")
    print(f"- Promedio compra hombres: ${df[df['Gender'] == 'Male']['Total'].mean():.2f}")
    
    print(f"\n2. ANÁLISIS POR HORARIO:")
    peak_hour = df.groupby('Hour')['Total'].sum().idxmax()
    peak_sales = df.groupby('Hour')['Total'].sum().max()
    print(f"- Hora pico de ventas: {peak_hour}:00 hrs (${peak_sales:,.2f})")
    
    morning_sales = df[df['Hour'].between(6, 11)]['Total'].sum()
    afternoon_sales = df[df['Hour'].between(12, 17)]['Total'].sum()
    evening_sales = df[df['Hour'].between(18, 23)]['Total'].sum()
    
    print(f"- Ventas mañana (6-11h): ${morning_sales:,.2f}")
    print(f"- Ventas tarde (12-17h): ${afternoon_sales:,.2f}")
    print(f"- Ventas noche (18-23h): ${evening_sales:,.2f}")
    
    print(f"\n3. ANÁLISIS POR LÍNEA DE PRODUCTOS:")
    top_product = df.groupby('Product line')['Total'].sum().idxmax()
    top_product_sales = df.groupby('Product line')['Total'].sum().max()
    print(f"- Línea más rentable: {top_product} (${top_product_sales:,.2f})")
    
    for product in df['Product line'].unique():
        product_sales = df[df['Product line'] == product]['Total'].sum()
        product_pct = (product_sales / df['Total'].sum()) * 100
        print(f"- {product}: ${product_sales:,.2f} ({product_pct:.1f}%)")
    
    print(f"\n4. ANÁLISIS POR SUCURSAL:")
    for branch in sorted(df['Branch'].unique()):
        branch_sales = df[df['Branch'] == branch]['Total'].sum()
        branch_pct = (branch_sales / df['Total'].sum()) * 100
        branch_transactions = len(df[df['Branch'] == branch])
        print(f"- Sucursal {branch}: ${branch_sales:,.2f} ({branch_pct:.1f}%) - {branch_transactions} transacciones")
    
    print(f"\n5. RECOMENDACIONES PARA CAMPAÑAS DE MARKETING:")
    print("- Enfocar campañas en las horas pico de ventas")
    print("- Desarrollar estrategias específicas por género según los patrones identificados")
    print("- Potenciar las líneas de productos más rentables")
    print("- Equilibrar el rendimiento entre sucursales")
    print("- Considerar promociones cruzadas entre líneas de productos complementarias")

def main():
    """Main analysis function"""
    # Load data
    df = load_and_clean_data('/Users/javierfernandez/Desktop/Taller/Libro1.csv')
    
    # Perform all analyses
    gender_stats = analyze_gender_purchases(df)
    hourly_stats = analyze_hourly_distribution(df)
    product_revenue = analyze_product_line_revenue(df)
    product_gender_revenue = analyze_product_line_by_gender(df)
    branch_stats = analyze_branch_purchases(df)
    
    # Generate summary report
    generate_summary_report(df, gender_stats, hourly_stats, product_revenue, product_gender_revenue, branch_stats)
    
    print(f"\nAnálisis completado. Gráficos guardados en:")
    print("- gender_analysis.png")
    print("- hourly_analysis.png") 
    print("- product_line_analysis.png")
    print("- product_gender_analysis.png")
    print("- branch_analysis.png")

if __name__ == "__main__":
    main()
