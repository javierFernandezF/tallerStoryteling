import pandas as pd

# Test different ways to load the data
file_path = '/Users/javierfernandez/Desktop/Taller/Libro1.csv'

print("Method 1: Default read_csv")
try:
    df1 = pd.read_csv(file_path)
    print(f"Columns: {df1.columns.tolist()}")
    print(f"Shape: {df1.shape}")
except Exception as e:
    print(f"Error: {e}")

print("\nMethod 2: With semicolon separator")
try:
    df2 = pd.read_csv(file_path, sep=';')
    print(f"Columns: {df2.columns.tolist()}")
    print(f"Shape: {df2.shape}")
except Exception as e:
    print(f"Error: {e}")

print("\nMethod 3: Manual parsing")
try:
    with open(file_path, 'r') as f:
        lines = f.readlines()[:3]
    
    print("First 3 lines:")
    for i, line in enumerate(lines):
        print(f"Line {i+1}: {repr(line)}")
    
    # Try to parse the header manually
    header_line = lines[0].strip()
    if header_line.startswith('"') and header_line.endswith('"'):
        header = header_line[1:-1].split(';')
        print(f"Parsed header: {header}")
        
        # Try to parse a data line
        data_line = lines[1].strip()
        if data_line.startswith('"') and data_line.endswith('"'):
            data = data_line[1:-1].split(';')
            print(f"Parsed data: {data[:5]}...")  # First 5 fields
            
except Exception as e:
    print(f"Error: {e}")
