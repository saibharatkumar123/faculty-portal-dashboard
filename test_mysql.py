import mysql.connector

try:
    print("🔍 Testing MySQL connection to Railway...")
    print("Host: crossover.proxy.rlwy.net")
    print("Port: 3306")
    print("User: root")
    
    conn = mysql.connector.connect(
        host='crossover.proxy.rlwy.net',
        user='root',
        password='tVTpsWGpAjrDUjkUnRbWHcuyUpHxlRWS',
        database='railway',
        port=3306,
        connect_timeout=10
    )
    print("✅ MySQL connection successful!")
    
    # Test a simple query
    cursor = conn.cursor()
    cursor.execute("SELECT 1 as test")
    result = cursor.fetchone()
    print(f"✅ Test query successful: {result}")
    
    conn.close()
    print("✅ Connection closed properly")
    
except Exception as e:
    print(f"❌ MySQL connection failed: {e}")
