import sqlite3

conn = sqlite3.connect('C:/AO20/dev/server/database.db')
cursor = conn.cursor()

# Query to reset everything faction-related
update_query = """
UPDATE user 
SET status = 1, 
    faction_score = 0, 
    ciudadanos_matados = 0, 
    criminales_matados = 0, 
    recompensas_real = 0, 
    recompensas_caos = 0, 
    recibio_armadura_real = 0, 
    recibio_armadura_caos = 0,
    reenlistadas = 0
WHERE lower(name) = 'talados'
"""

cursor.execute(update_query)
conn.commit()

# Verify
cursor.execute("SELECT id, name, status, faction_score, ciudadanos_matados FROM user WHERE lower(name) = 'talados'")
row = cursor.fetchone()
print(f"Verified: {row}")

conn.close()
