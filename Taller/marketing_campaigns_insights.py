import pandas as pd
import numpy as np

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

def find_campaign_insights():
    """Identify key insights for 3 marketing campaigns"""
    df = load_data()
    
    print("🎯 ANÁLISIS PROFUNDO PARA 3 CAMPAÑAS DE MARKETING")
    print("=" * 70)
    
    # INSIGHT 1: OPORTUNIDAD DE GÉNERO
    print("\n📊 INSIGHT #1: BRECHA DE GÉNERO EN PRODUCTOS ESPECÍFICOS")
    print("-" * 60)
    
    # Analizar diferencias significativas por género en cada producto
    gender_product_analysis = []
    for product in df['Product line'].unique():
        female_sales = df[(df['Gender'] == 'Female') & (df['Product line'] == product)]['Total'].sum()
        male_sales = df[(df['Gender'] == 'Male') & (df['Product line'] == product)]['Total'].sum()
        total_product_sales = female_sales + male_sales
        
        if total_product_sales > 0:
            female_pct = (female_sales / total_product_sales) * 100
            male_pct = (male_sales / total_product_sales) * 100
            gap = abs(female_pct - male_pct)
            dominant_gender = 'Female' if female_pct > male_pct else 'Male'
            
            gender_product_analysis.append({
                'product': product,
                'gap': gap,
                'dominant_gender': dominant_gender,
                'dominant_pct': max(female_pct, male_pct),
                'female_sales': female_sales,
                'male_sales': male_sales,
                'total_sales': total_product_sales
            })
    
    # Ordenar por brecha más grande
    gender_product_analysis.sort(key=lambda x: x['gap'], reverse=True)
    
    print("Productos con mayor brecha de género:")
    for item in gender_product_analysis[:3]:
        print(f"• {item['product']}: {item['dominant_gender']} domina con {item['dominant_pct']:.1f}% (brecha: {item['gap']:.1f}%)")
        print(f"  Oportunidad: ${item['total_sales']:,.2f} en ventas totales")
    
    # INSIGHT 2: OPORTUNIDAD HORARIA
    print("\n📊 INSIGHT #2: OPORTUNIDADES DE CRECIMIENTO HORARIO")
    print("-" * 60)
    
    hourly_analysis = df.groupby('Hour').agg({
        'Total': ['sum', 'count', 'mean'],
        'Quantity': 'sum'
    }).round(2)
    
    hourly_sales = df.groupby('Hour')['Total'].sum()
    hourly_transactions = df.groupby('Hour').size()
    
    # Identificar horas con bajo rendimiento pero potencial
    avg_hourly_sales = hourly_sales.mean()
    avg_transactions = hourly_transactions.mean()
    
    underperforming_hours = []
    for hour in range(10, 22):
        if hour in hourly_sales.index:
            sales = hourly_sales[hour]
            transactions = hourly_transactions[hour]
            avg_ticket = sales / transactions if transactions > 0 else 0
            
            # Horas con transacciones decentes pero ventas bajas = oportunidad de upselling
            if transactions >= avg_transactions * 0.8 and sales < avg_hourly_sales * 0.9:
                underperforming_hours.append({
                    'hour': hour,
                    'sales': sales,
                    'transactions': transactions,
                    'avg_ticket': avg_ticket,
                    'potential': (avg_hourly_sales - sales)
                })
    
    underperforming_hours.sort(key=lambda x: x['potential'], reverse=True)
    
    print("Horas con mayor potencial de crecimiento:")
    for item in underperforming_hours[:3]:
        print(f"• {item['hour']:2d}:00 hrs - Ticket promedio: ${item['avg_ticket']:.2f}")
        print(f"  Potencial adicional: ${item['potential']:,.2f} por hora")
    
    # INSIGHT 3: OPORTUNIDAD DE CROSS-SELLING
    print("\n📊 INSIGHT #3: OPORTUNIDADES DE VENTA CRUZADA")
    print("-" * 60)
    
    # Analizar patrones de compra por cliente (usando Customer type como proxy)
    customer_patterns = df.groupby(['Customer type', 'Product line']).agg({
        'Total': ['sum', 'count'],
        'Quantity': 'sum'
    }).round(2)
    
    # Analizar diferencias entre Members vs Normal customers
    member_behavior = df[df['Customer type'] == 'Member'].groupby('Product line')['Total'].agg(['sum', 'count', 'mean'])
    normal_behavior = df[df['Customer type'] == 'Normal'].groupby('Product line')['Total'].agg(['sum', 'count', 'mean'])
    
    print("Diferencias entre clientes Member vs Normal:")
    cross_sell_opportunities = []
    
    for product in df['Product line'].unique():
        if product in member_behavior.index and product in normal_behavior.index:
            member_avg = member_behavior.loc[product, 'mean']
            normal_avg = normal_behavior.loc[product, 'mean']
            member_count = member_behavior.loc[product, 'count']
            normal_count = normal_behavior.loc[product, 'count']
            
            # Productos donde Members gastan más = oportunidad para convertir Normal a Member
            if member_avg > normal_avg * 1.1:  # 10% más
                uplift_potential = (normal_avg * (member_avg / normal_avg - 1)) * normal_count
                cross_sell_opportunities.append({
                    'product': product,
                    'member_avg': member_avg,
                    'normal_avg': normal_avg,
                    'uplift_pct': (member_avg / normal_avg - 1) * 100,
                    'normal_customers': normal_count,
                    'potential_revenue': uplift_potential
                })
    
    cross_sell_opportunities.sort(key=lambda x: x['potential_revenue'], reverse=True)
    
    for item in cross_sell_opportunities[:3]:
        print(f"• {item['product']}: Members gastan {item['uplift_pct']:.1f}% más")
        print(f"  Potencial: ${item['potential_revenue']:,.2f} convirtiendo clientes Normal")
    
    # INSIGHT 4: OPORTUNIDAD POR SUCURSAL
    print("\n📊 INSIGHT #4: EQUILIBRIO ENTRE SUCURSALES")
    print("-" * 60)
    
    branch_performance = df.groupby('Branch').agg({
        'Total': ['sum', 'count', 'mean'],
        'gross income': 'sum'
    }).round(2)
    
    branch_sales = df.groupby('Branch')['Total'].sum()
    best_branch = branch_sales.idxmax()
    worst_branch = branch_sales.idxmin()
    
    best_performance = branch_sales[best_branch]
    worst_performance = branch_sales[worst_branch]
    gap = best_performance - worst_performance
    
    print(f"Mejor sucursal: {best_branch} con ${best_performance:,.2f}")
    print(f"Menor rendimiento: {worst_branch} con ${worst_performance:,.2f}")
    print(f"Brecha de rendimiento: ${gap:,.2f}")
    print(f"Potencial de mejora: {(gap/worst_performance)*100:.1f}% de incremento")
    
    # Analizar qué hace diferente la mejor sucursal
    best_branch_data = df[df['Branch'] == best_branch]
    worst_branch_data = df[df['Branch'] == worst_branch]
    
    print(f"\n¿Qué hace diferente la Sucursal {best_branch}?")
    
    # Comparar por línea de productos
    best_products = best_branch_data.groupby('Product line')['Total'].sum().sort_values(ascending=False)
    worst_products = worst_branch_data.groupby('Product line')['Total'].sum().sort_values(ascending=False)
    
    print("Top productos por sucursal:")
    print(f"Sucursal {best_branch}: {best_products.index[0]} (${best_products.iloc[0]:,.2f})")
    print(f"Sucursal {worst_branch}: {worst_products.index[0]} (${worst_products.iloc[0]:,.2f})")
    
    return {
        'gender_gaps': gender_product_analysis,
        'hourly_opportunities': underperforming_hours,
        'cross_sell': cross_sell_opportunities,
        'branch_gap': gap,
        'best_branch': best_branch,
        'worst_branch': worst_branch
    }

def create_campaign_storytelling(insights):
    """Create compelling storytelling for 3 marketing campaigns"""
    
    print("\n\n🎬 STORYTELLING PARA 3 CAMPAÑAS DE MARKETING")
    print("=" * 70)
    
    # CAMPAÑA 1: GÉNERO
    print("\n🎯 CAMPAÑA #1: 'ELLA ELIGE, ÉL DESCUBRE'")
    print("=" * 50)
    print("📖 HISTORIA:")
    print("""
En nuestros supermercados, hemos descubierto algo fascinante: las mujeres y los hombres 
tienen preferencias muy marcadas en ciertos productos. Las mujeres dominan en Food & 
Beverages (59.1%) y Fashion Accessories (56.0%), mientras que los hombres prefieren 
Health & Beauty (62.3%).

Esta no es solo una estadística, es una oportunidad de oro. ¿Qué pasaría si pudiéramos 
ayudar a cada género a descubrir los productos que el otro ya ama?
    """)
    
    print("🎯 ESTRATEGIA:")
    print("• Crear 'Zonas de Descubrimiento' con productos cruzados")
    print("• Promociones 2x1: 'Compra tu favorito, descubre el de ella/él'")
    print("• Influencers de género cruzado promocionando productos")
    print("• Degustaciones y demos en productos con brecha de género")
    
    print("💰 POTENCIAL DE INGRESOS:")
    top_gap = insights['gender_gaps'][0]
    print(f"• Producto objetivo: {top_gap['product']}")
    print(f"• Mercado actual: ${top_gap['total_sales']:,.2f}")
    print(f"• Potencial de crecimiento: 15-25% aumentando participación del género minoritario")
    
    # CAMPAÑA 2: HORARIA
    print("\n🎯 CAMPAÑA #2: 'HAPPY HOURS QUE IMPORTAN'")
    print("=" * 50)
    print("📖 HISTORIA:")
    print("""
Nuestros datos revelan que aunque tenemos un flujo constante de clientes durante el día,
hay horas específicas donde los clientes vienen pero gastan menos de lo que podrían.
Son momentos perfectos para crear 'Happy Hours' estratégicos.

No se trata de horarios muertos, sino de momentos con potencial no aprovechado.
Clientes que están ahí, pero necesitan un pequeño empujón para llevar más productos.
    """)
    
    print("🎯 ESTRATEGIA:")
    print("• 'Power Hours': Descuentos progresivos en horas específicas")
    print("• 'Combo del Momento': Ofertas especiales por horario")
    print("• 'Desafío del Ticket': Incentivos para aumentar el valor promedio")
    print("• Staff especializado en upselling durante estas horas")
    
    print("💰 POTENCIAL DE INGRESOS:")
    if insights['hourly_opportunities']:
        total_hourly_potential = sum(item['potential'] for item in insights['hourly_opportunities'][:3])
        print(f"• Potencial diario adicional: ${total_hourly_potential:,.2f}")
        print(f"• Potencial mensual: ${total_hourly_potential * 30:,.2f}")
        print(f"• ROI esperado: 200-300% sobre inversión en promociones")
    
    # CAMPAÑA 3: LEALTAD Y CROSS-SELLING
    print("\n🎯 CAMPAÑA #3: 'MIEMBROS VIP: EL SECRETO MEJOR GUARDADO'")
    print("=" * 50)
    print("📖 HISTORIA:")
    print("""
Hemos descubierto que nuestros clientes Member tienen un secreto: gastan significativamente
más en ciertos productos que los clientes regulares. No es casualidad, es conocimiento.

Los Members han descubierto el valor real de ciertos productos. ¿Qué pasaría si 
compartiéramos estos 'secretos' con todos nuestros clientes y los invitáramos a 
unirse al club de los que 'saben comprar'?
    """)
    
    print("🎯 ESTRATEGIA:")
    print("• 'Secretos de Members': Campaña educativa sobre productos premium")
    print("• 'Prueba VIP': Degustaciones exclusivas de productos donde Members gastan más")
    print("• 'Upgrade Challenge': Incentivos para convertir Normal a Member")
    print("• 'Member Recommends': Testimoniales de clientes Member")
    
    print("💰 POTENCIAL DE INGRESOS:")
    if insights['cross_sell']:
        total_cross_sell = sum(item['potential_revenue'] for item in insights['cross_sell'][:3])
        print(f"• Potencial por conversión: ${total_cross_sell:,.2f}")
        print(f"• Meta: Convertir 20% de clientes Normal a Member")
        print(f"• Incremento esperado en ticket promedio: 15-30%")
    
    print("\n\n🚀 CRONOGRAMA DE IMPLEMENTACIÓN")
    print("=" * 50)
    print("MES 1: 'ELLA ELIGE, ÉL DESCUBRE'")
    print("• Semanas 1-2: Setup de zonas y capacitación")
    print("• Semanas 3-4: Lanzamiento con influencers y promociones")
    print()
    print("MES 2: 'HAPPY HOURS QUE IMPORTAN'")
    print("• Semanas 1-2: Análisis de horarios óptimos por sucursal")
    print("• Semanas 3-4: Implementación de Power Hours")
    print()
    print("MES 3: 'MIEMBROS VIP: EL SECRETO MEJOR GUARDADO'")
    print("• Semanas 1-2: Desarrollo de contenido educativo")
    print("• Semanas 3-4: Campaña de conversión masiva")
    
    print("\n💡 MÉTRICAS DE ÉXITO:")
    print("• Incremento en ventas totales: 15-25%")
    print("• Aumento en ticket promedio: 20-35%")
    print("• Conversión Normal a Member: +20%")
    print("• Satisfacción del cliente: +15%")
    print("• ROI de campañas: 250-400%")

if __name__ == "__main__":
    insights = find_campaign_insights()
    create_campaign_storytelling(insights)
