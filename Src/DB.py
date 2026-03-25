import sqlite3
conn = sqlite3.connect("products.db")
cursor = conn.cursor()
def Dumping():
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS dumping (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        model TEXT,
        brand TEXT,
        type TEXT,
        status TEXT)
                
    """)
    conn.commit()

# def deleteing():
#     cursor.execute("DELETE FROM dumping")
#     conn.commit()
# def adding_colum():
#     cursor.execute("ALTER TABLE dumping ADD COLUMN description TEXT")
#     conn.commit()
#     print("successfull")
    


# def Table():
#     cursor.execute("""
#     CREATE TABLE IF NOT EXISTS products (
#         id INTEGER PRIMARY KEY AUTOINCREMENT,
#         brand TEXT,
#         model TEXT,
#         type TEXT,
#         description TEXT)
#     """)
#     conn.commit()
# def deleteing():
#     cursor.execute("DELETE FROM products")
#     conn.commit()


def inserting_initial(data):
    cursor.executemany("""
    INSERT INTO dumping (model, brand, type , status)
    VALUES (?, ?, ? , ? )
    """, data)
    conn.commit()
def indexation():
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_lookup ON dumping(model, brand, type, status)")
    conn.commit()
def inserting_descriptions(batch_data):
    cursor.executemany("""
    UPDATE dumping SET description = ? ,status = ? WHERE model = ? and brand=? and type=? and status= "Pending"
    """, (batch_data))
    conn.commit()








# def inserting2(data):
#     cursor.executemany("""
#     INSERT INTO products (brand, model, Type , description)
#     VALUES (?, ?, ? , ? )
#     """, data)
#     conn.commit()


def fetching_data1():
    # cursor.execute("SELECT model, brand, type , status FROM dumping where brand = 'LOGITECH' OR brand = 'VEEAM SUPPORT'")
    cursor.execute("SELECT model, brand, type, status FROM dumping WHERE status = 'Pending' ORDER BY id LIMIT 200000")
    return cursor.fetchall()

def fetching_data2():
    # cursor.execute("SELECT model, brand, type , status FROM dumping where brand = 'LOGITECH' OR brand = 'VEEAM SUPPORT'")
    cursor.execute("SELECT model, brand, type , status , description FROM dumping where status = 'Not_Pending' ORDER BY id")


    output = []
    while True:
        rows = cursor.fetchall()
        if not rows:
            break
        for row in rows:
            output.append(row)
    
    return output


















# def fetching_data(batch_size = 1000):
#     cursor.execute("SELECT brand, model, Type , description FROM products")


#     output = []
#     while True:
#         rows = cursor.fetchmany(batch_size)
#         if not rows:
#             break
#         for row in rows:
#             output.append(row)
#             print(output)
#     return output
