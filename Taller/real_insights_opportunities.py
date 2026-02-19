import pandas as pd

def load_data():
    """Load and clean the supermarket sales data"""
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
    numeric_cols = ['Unit price', 'Quantity', 'Tax 5%', 'Total', 'Time', 'cogs', 'gross margin percentage', 'gross income', 'Rating']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].str.replace(',', '.').astype(float)
    
    df['Hour'] = (df['Time'] * 24).round().astype(int).clip(0, 23)
    return df

def find_real_opportunities():
    """Find 3 concrete opportunities based on actual data patterns"""
    df = load_data()
    
    print("🔍 ANÁLISIS DE OPORTUNIDADES REALES BASADAS EN DATOS")
    print("=" * 65)
    
    print("\n📊 DATOS GENERALES:")
    print(f"• Total transacciones: {len(df):,}")
    print(f"• Total ventas: ${df['Total'].sum():,.2f}")
    print(f"• Ticket promedio: ${df['Total'].mean():.2f}")
    
    # OPORTUNIDAD 1: DESEQUILIBRIO DE GÉNERO EN PRODUCTOS ESPECÍFICOS
    print("\n🎯 OPORTUNIDAD #1: DESEQUILIBRIO DE GÉNERO EN PRODUCTOS")
    print("-" * 60)
    print("LO QUE LOS DATOS REVELAN:")
    
    # Analizar cada producto por género
    gender_analysis = {}
    for product in df['Product line'].unique():
        product_data = df[df['Product line'] == product]
        female_sales = product_data[product_data['Gender'] == 'Female']['Total'].sum()
        male_sales = product_data[product_data['Gender'] == 'Male']['Total'].sum()
        total_sales = female_sales + male_sales
        
        female_count = len(product_data[product_data['Gender'] == 'Female'])
        male_count = len(product_data[product_data['Gender'] == 'Male'])
        total_count = female_count + male_count
        
        if total_sales > 0:
            female_pct = (female_sales / total_sales) * 100
            male_pct = (male_sales / total_sales) * 100
            
            gender_analysis[product] = {
                'female_pct': female_pct,
                'male_pct': male_pct,
                'female_sales': female_sales,
                'male_sales': male_sales,
                'total_sales': total_sales,
                'female_count': female_count,
                'male_count': male_count,
                'gap': abs(female_pct - male_pct)
            }
    
    # Mostrar los 3 productos con mayor desequilibrio
    sorted_products = sorted(gender_analysis.items(), key=lambda x: x[1]['gap'], reverse=True)
    
    for i, (product, data) in enumerate(sorted_products[:3], 1):
        dominant_gender = 'Mujeres' if data['female_pct'] > data['male_pct'] else 'Hombres'
        minority_gender = 'Hombres' if data['female_pct'] > data['male_pct'] else 'Mujeres'
        dominant_pct = max(data['female_pct'], data['male_pct'])
        minority_pct = min(data['female_pct'], data['male_pct'])
        
        print(f"{i}. {product}:")
        print(f"   • {dominant_gender}: {dominant_pct:.1f}% (${max(data['female_sales'], data['male_sales']):,.2f})")
        print(f"   • {minority_gender}: {minority_pct:.1f}% (${min(data['female_sales'], data['male_sales']):,.2f})")
        print(f"   • Brecha: {data['gap']:.1f}% - Mercado total: ${data['total_sales']:,.2f}")
        print(f"   • OPORTUNIDAD: El género minoritario está subrepresentado")
    
    # OPORTUNIDAD 2: DIFERENCIAS EN TICKET PROMEDIO POR TIPO DE CLIENTE
    print("\n🎯 OPORTUNIDAD #2: DIFERENCIAS EN COMPORTAMIENTO DE CLIENTES")
    print("-" * 60)
    print("LO QUE LOS DATOS REVELAN:")
    
    # Comparar Members vs Normal
    member_data = df[df['Customer type'] == 'Member']
    normal_data = df[df['Customer type'] == 'Normal']
    
    member_avg_ticket = member_data['Total'].mean()
    normal_avg_ticket = normal_data['Total'].mean()
    member_total_sales = member_data['Total'].sum()
    normal_total_sales = normal_data['Total'].sum()
    
    print(f"CLIENTES MEMBER:")
    print(f"• Cantidad: {len(member_data)} ({len(member_data)/len(df)*100:.1f}%)")
    print(f"• Ticket promedio: ${member_avg_ticket:.2f}")
    print(f"• Ventas totales: ${member_total_sales:,.2f}")
    
    print(f"\nCLIENTES NORMAL:")
    print(f"• Cantidad: {len(normal_data)} ({len(normal_data)/len(df)*100:.1f}%)")
    print(f"• Ticket promedio: ${normal_avg_ticket:.2f}")
    print(f"• Ventas totales: ${normal_total_sales:,.2f}")
    
    ticket_difference = member_avg_ticket - normal_avg_ticket
    print(f"\nDIFERENCIA:")
    print(f"• Members gastan ${ticket_difference:.2f} más por transacción ({(member_avg_ticket/normal_avg_ticket-1)*100:.1f}% más)")
    print(f"• OPORTUNIDAD: {len(normal_data)} clientes Normal podrían convertirse a Member")
    
    # Analizar por producto donde la diferencia es mayor
    print(f"\nPRODUCTOS DONDE MEMBERS GASTAN MÁS:")
    member_product_avg = member_data.groupby('Product line')['Total'].mean()
    normal_product_avg = normal_data.groupby('Product line')['Total'].mean()
    
    product_differences = []
    for product in df['Product line'].unique():
        if product in member_product_avg.index and product in normal_product_avg.index:
            member_avg = member_product_avg[product]
            normal_avg = normal_product_avg[product]
            difference = member_avg - normal_avg
            pct_diff = (member_avg / normal_avg - 1) * 100
            
            product_differences.append({
                'product': product,
                'member_avg': member_avg,
                'normal_avg': normal_avg,
                'difference': difference,
                'pct_diff': pct_diff
            })
    
    product_differences.sort(key=lambda x: x['pct_diff'], reverse=True)
    
    for i, item in enumerate(product_differences[:3], 1):
        print(f"{i}. {item['product']}: Members ${item['member_avg']:.2f} vs Normal ${item['normal_avg']:.2f}")
        print(f"   • Diferencia: ${item['difference']:.2f} ({item['pct_diff']:.1f}% más)")
    
    # OPORTUNIDAD 3: HORARIOS CON MENOR DENSIDAD DE VENTAS
    print("\n🎯 OPORTUNIDAD #3: HORARIOS CON BAJA DENSIDAD DE VENTAS")
    print("-" * 60)
    print("LO QUE LOS DATOS REVELAN:")
    
    # Analizar ventas por hora
    hourly_stats = df.groupby('Hour').agg({
        'Total': ['sum', 'count', 'mean']
    }).round(2)
    
    hourly_sales = df.groupby('Hour')['Total'].sum()
    hourly_count = df.groupby('Hour').size()
    hourly_avg_ticket = hourly_sales / hourly_count
    
    print("ANÁLISIS POR HORA:")
    print("Hora | Ventas    | Trans | Ticket Prom")
    print("-" * 40)
    
    hour_performance = []
    for hour in range(10, 22):
        if hour in hourly_sales.index:
            sales = hourly_sales[hour]
            count = hourly_count[hour]
            avg_ticket = hourly_avg_ticket[hour]
            
            hour_performance.append({
                'hour': hour,
                'sales': sales,
                'count': count,
                'avg_ticket': avg_ticket
            })
            
            print(f"{hour:2d}h  | ${sales:8,.0f} | {count:5d} | ${avg_ticket:7.2f}")
    
    # Identificar horas con menor rendimiento
    avg_sales_per_hour = sum(item['sales'] for item in hour_performance) / len(hour_performance)
    avg_ticket_overall = df['Total'].mean()
    
    print(f"\nPROMEDIOS GENERALES:")
    print(f"• Ventas promedio por hora: ${avg_sales_per_hour:,.2f}")
    print(f"• Ticket promedio general: ${avg_ticket_overall:.2f}")
    
    underperforming_hours = []
    for item in hour_performance:
        if item['sales'] < avg_sales_per_hour * 0.9:  # 10% por debajo del promedio
            underperforming_hours.append(item)
    
    print(f"\nHORAS CON VENTAS POR DEBAJO DEL PROMEDIO:")
    for item in underperforming_hours:
        gap = avg_sales_per_hour - item['sales']
        print(f"• {item['hour']:2d}:00h - ${item['sales']:,.0f} (${gap:,.0f} menos que el promedio)")
        print(f"  Transacciones: {item['count']} | Ticket: ${item['avg_ticket']:.2f}")
    
    print(f"\nOPORTUNIDAD: {len(underperforming_hours)} horas tienen ventas por debajo del promedio")
    
    # RESUMEN DE OPORTUNIDADES
    print("\n" + "="*65)
    print("📋 RESUMEN DE LAS 3 OPORTUNIDADES IDENTIFICADAS")
    print("="*65)
    
    print("\n1. DESEQUILIBRIO DE GÉNERO:")
    top_gender_gap = sorted_products[0]
    print(f"   • Producto: {top_gender_gap[0]}")
    print(f"   • Brecha: {top_gender_gap[1]['gap']:.1f}%")
    print(f"   • Mercado: ${top_gender_gap[1]['total_sales']:,.2f}")
    print(f"   • ACCIÓN: Atraer al género minoritario a este producto")
    
    print(f"\n2. CONVERSIÓN DE CLIENTES:")
    print(f"   • {len(normal_data)} clientes Normal vs {len(member_data)} Members")
    print(f"   • Diferencia en ticket: ${ticket_difference:.2f} ({(member_avg_ticket/normal_avg_ticket-1)*100:.1f}% más)")
    print(f"   • ACCIÓN: Convertir clientes Normal a Member")
    
    print(f"\n3. OPTIMIZACIÓN HORARIA:")
    total_gap = sum(avg_sales_per_hour - item['sales'] for item in underperforming_hours)
    print(f"   • {len(underperforming_hours)} horas por debajo del promedio")
    print(f"   • Brecha total: ${total_gap:,.2f} por día")
    print(f"   • ACCIÓN: Aumentar ventas en horarios de menor rendimiento")

if __name__ == "__main__":
    find_real_opportunities()
